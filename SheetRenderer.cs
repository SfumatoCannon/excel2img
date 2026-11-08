using SkiaSharp;

namespace Excel2Img;

internal sealed class SheetRenderer : IDisposable
{
    private const float Margin = 10f;
    private const float TextPadding = 1f;
    private const float GridLineWidth = 1f;

    private static readonly SKColor GridLineColor = new(0xD9, 0xD9, 0xD9);

    private readonly Workbook _workbook;
    private readonly SheetInfo _sheet;
    private readonly double _scale;
    private readonly bool _showGridLines;

    private readonly Dictionary<int, double> _colWidths = new();
    private readonly Dictionary<int, double> _rowHeights = new();
    private double[] _colX = Array.Empty<double>();
    private double[] _rowY = Array.Empty<double>();
    private readonly Dictionary<(int Row, int Col), MergedRange> _mergeAnchors = new();
    private readonly HashSet<(int Row, int Col)> _mergedCells = new();
    private readonly Dictionary<uint, SKPaint> _fillPaints = new();
    private readonly List<ImagePlacement> _imagePlacements = new();
    private readonly int _maxDigitWidth;

    private double _contentWidth;
    private double _contentHeight;

    public SheetRenderer(Workbook workbook, SheetInfo sheet, double scale, bool showGridLines)
    {
        _workbook = workbook;
        _sheet = sheet;
        _scale = scale;
        _showGridLines = showGridLines;
        _maxDigitWidth = MeasureMaxDigitWidth(workbook);

        foreach (var range in sheet.Merged)
        {
            _mergeAnchors[(range.FirstRow, range.FirstCol)] = range;
            for (var row = range.FirstRow; row <= range.LastRow; row++)
            {
                for (var col = range.FirstCol; col <= range.LastCol; col++)
                {
                    _mergedCells.Add((row, col));
                }
            }
        }

        BuildLayout();
        BuildImageLayout();
    }

    public (int Width, int Height) RenderTo(string outputPath)
    {
        using var bitmap = Render();
        var ext = Path.GetExtension(outputPath).ToLowerInvariant();
        var format = ext switch
        {
            ".png" => SKEncodedImageFormat.Png,
            ".jpg" or ".jpeg" => SKEncodedImageFormat.Jpeg,
            _ => throw new CliException(Localization.T(
                "Output file extension must be .png, .jpg or .jpeg.",
                "输出文件扩展名必须是 .png、.jpg 或 .jpeg。")),
        };

        var directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using var image = SKImage.FromBitmap(bitmap);
        using var data = image.Encode(format, 95)
            ?? throw new CliException(Localization.T(
                "Failed to encode the image.",
                "图像编码失败。"));
        using (var stream = File.Create(outputPath))
        {
            data.SaveTo(stream);
        }

        return (bitmap.Width, bitmap.Height);
    }

    private SKBitmap Render()
    {
        var width = (int)Math.Ceiling((_contentWidth + Margin * 2) * _scale);
        var height = (int)Math.Ceiling((_contentHeight + Margin * 2) * _scale);
        var info = new SKImageInfo(Math.Max(1, width), Math.Max(1, height), SKColorType.Bgra8888, SKAlphaType.Premul);
        var bitmap = new SKBitmap(info);

        using var canvas = new SKCanvas(bitmap);
        canvas.Clear(SKColors.White);
        canvas.Scale((float)_scale);
        canvas.Translate(Margin, Margin);

        if (!_sheet.IsEmpty)
        {
            DrawSheet(canvas);
        }

        DrawImages(canvas);

        canvas.Flush();
        return bitmap;
    }

    private void DrawSheet(SKCanvas canvas)
    {
        using var gridPaint = new SKPaint
        {
            Style = SKPaintStyle.Stroke,
            StrokeWidth = GridLineWidth,
            Color = GridLineColor,
            IsAntialias = false,
        };

        using var borderPaint = new SKPaint
        {
            Style = SKPaintStyle.Stroke,
            IsAntialias = true,
        };

        for (var row = _sheet.MinRow; row <= _sheet.MaxRow; row++)
        {
            for (var col = _sheet.MinCol; col <= _sheet.MaxCol; col++)
            {
                var key = (row, col);
                if (_mergedCells.Contains(key) && !_mergeAnchors.ContainsKey(key))
                {
                    continue;
                }

                _sheet.Cells.TryGetValue(key, out var cell);
                var style = GetStyle(cell?.StyleIndex ?? 0);
                var rect = GetCellRect(row, col);

                if (rect.Width <= 0 || rect.Height <= 0)
                {
                    continue;
                }

                if (style.FillColor is { } fillColor)
                {
                    canvas.DrawRect(rect, GetFillPaint(fillColor));
                }

                var conditional = _sheet.ConditionalFills.TryGetValue(key, out var conditionalColor);
                if (conditional)
                {
                    canvas.DrawRect(rect, GetFillPaint(conditionalColor));
                }

                if (_showGridLines && style.FillColor == null && !conditional)
                {
                    canvas.DrawRect(rect, gridPaint);
                }

                var top = style.Top;
                var right = style.Right;
                var bottom = style.Bottom;
                var left = style.Left;

                // 合并单元格四周的边框可能分别保存在边缘的单元格上
                // （例如竖直合并时，底边框通常放在最后一行的单元格上），
                // 因此按合并区域的四角单元格组合出外边框。
                if (_mergeAnchors.TryGetValue(key, out var merge))
                {
                    if (merge.LastRow != merge.FirstRow
                        && GetStyleAt(merge.LastRow, merge.FirstCol) is { Bottom.Visible: true } bottomStyle)
                    {
                        bottom = bottomStyle.Bottom;
                    }

                    if (merge.LastCol != merge.FirstCol
                        && GetStyleAt(merge.FirstRow, merge.LastCol) is { Right.Visible: true } rightStyle)
                    {
                        right = rightStyle.Right;
                    }
                }

                DrawBorders(canvas, borderPaint, rect, top, right, bottom, left);

                if (cell != null)
                {
                    DrawCellText(canvas, rect, cell, style);
                }
            }
        }
    }

    private void DrawImages(SKCanvas canvas)
    {
        foreach (var placement in _imagePlacements)
        {
            SKBitmap? bitmap;
            try
            {
                bitmap = SKBitmap.Decode(placement.Image.Data);
            }
            catch (Exception)
            {
                continue;
            }

            if (bitmap == null)
            {
                continue;
            }

            using (bitmap)
            {
                using var paint = new SKPaint
                {
                    IsAntialias = true,
                    FilterQuality = SKFilterQuality.High,
                };

                var source = GetImageSourceRect(bitmap, placement.Image);
                var destination = placement.Dest;
                var image = placement.Image;

                if (image.RotationDegrees % 360 != 0 || image.FlipHorizontal || image.FlipVertical)
                {
                    canvas.Save();
                    canvas.Translate(destination.MidX, destination.MidY);
                    if (image.RotationDegrees != 0)
                    {
                        canvas.RotateDegrees((float)image.RotationDegrees);
                    }

                    canvas.Scale(image.FlipHorizontal ? -1 : 1, image.FlipVertical ? -1 : 1);
                    var local = SKRect.Create(
                        -destination.Width / 2f,
                        -destination.Height / 2f,
                        destination.Width,
                        destination.Height);
                    canvas.DrawBitmap(bitmap, source, local, paint);
                    canvas.Restore();
                }
                else
                {
                    canvas.DrawBitmap(bitmap, source, destination, paint);
                }
            }
        }
    }

    private static SKRect GetImageSourceRect(SKBitmap bitmap, SheetImage image)
    {
        var left = Math.Clamp(image.CropLeft, 0.0, 0.99);
        var top = Math.Clamp(image.CropTop, 0.0, 0.99);
        var right = Math.Clamp(image.CropRight, 0.0, 0.99);
        var bottom = Math.Clamp(image.CropBottom, 0.0, 0.99);

        if (left + right >= 0.99 || top + bottom >= 0.99)
        {
            return new SKRect(0, 0, bitmap.Width, bitmap.Height);
        }

        return new SKRect(
            (float)(left * bitmap.Width),
            (float)(top * bitmap.Height),
            (float)((1 - right) * bitmap.Width),
            (float)((1 - bottom) * bitmap.Height));
    }

    private void BuildImageLayout()
    {
        const double emuPerPixel = 9525.0;

        foreach (var image in _sheet.Images)
        {
            double left, top, right, bottom;
            switch (image.Anchor)
            {
                case ImageAnchorKind.TwoCell:
                    left = ColumnLeft(image.FromCol) + image.FromColOffset / emuPerPixel;
                    top = RowTop(image.FromRow) + image.FromRowOffset / emuPerPixel;
                    right = ColumnLeft(image.ToCol) + image.ToColOffset / emuPerPixel;
                    bottom = RowTop(image.ToRow) + image.ToRowOffset / emuPerPixel;
                    break;
                case ImageAnchorKind.OneCell:
                    left = ColumnLeft(image.FromCol) + image.FromColOffset / emuPerPixel;
                    top = RowTop(image.FromRow) + image.FromRowOffset / emuPerPixel;
                    right = left + image.ExtCx / emuPerPixel;
                    bottom = top + image.ExtCy / emuPerPixel;
                    break;
                default:
                    left = image.PosX / emuPerPixel;
                    top = image.PosY / emuPerPixel;
                    right = left + image.ExtCx / emuPerPixel;
                    bottom = top + image.ExtCy / emuPerPixel;
                    break;
            }

            if (right - left < 1 || bottom - top < 1)
            {
                continue;
            }

            var rect = new SKRect((float)left, (float)top, (float)right, (float)bottom);
            _imagePlacements.Add(new ImagePlacement(image, rect));
            _contentWidth = Math.Max(_contentWidth, right);
            _contentHeight = Math.Max(_contentHeight, bottom);
        }
    }

    private double ColumnLeft(int col) =>
        col >= 0 && col < _colX.Length ? _colX[col] : 0;

    private double RowTop(int row) =>
        row >= 0 && row < _rowY.Length ? _rowY[row] : 0;

    private sealed record ImagePlacement(SheetImage Image, SKRect Dest);

    private SKRect GetCellRect(int row, int col)
    {
        if (_mergeAnchors.TryGetValue((row, col), out var merge))
        {
            var left = (float)_colX[merge.FirstCol];
            var top = (float)_rowY[merge.FirstRow];
            var right = (float)(_colX[merge.LastCol] + _colWidths[merge.LastCol]);
            var bottom = (float)(_rowY[merge.LastRow] + _rowHeights[merge.LastRow]);
            return new SKRect(left, top, right, bottom);
        }

        return new SKRect(
            (float)_colX[col],
            (float)_rowY[row],
            (float)(_colX[col] + _colWidths[col]),
            (float)(_rowY[row] + _rowHeights[row]));
    }

    private void DrawCellText(SKCanvas canvas, SKRect rect, CellData cell, CellStyle style)
    {
        var text = NumberFormatter.FormatValue(cell, style, _workbook);
        if (string.IsNullOrEmpty(text))
        {
            return;
        }

        var fontPx = (float)(style.FontSizePt * 4.0 / 3.0);
        if (fontPx <= 0)
        {
            return;
        }

        using var paint = CreateTextPaint(style, fontPx, text);
        var fontMetrics = paint.FontMetrics;
        var lineHeight = GetLineHeight(style);
        var indent = style.Indent * 9f;
        var availableWidth = rect.Width - TextPadding * 2 - indent;
        if (availableWidth <= 1)
        {
            return;
        }

        var lines = style.WrapText
            ? WrapText(text, availableWidth, paint)
            : SplitLines(text);
        if (lines.Count == 0)
        {
            return;
        }

        var textHeight = -fontMetrics.Ascent + fontMetrics.Descent;

        // 换行内容略高于单元格时（不同字体/换行算法的细微差异），
        // 适当压缩行距以保证内容完整显示，而不是在底部裁掉一行。
        if (lines.Count > 1)
        {
            var needed = (lines.Count - 1) * lineHeight + textHeight;
            var available = rect.Height - TextPadding * 2 - 10;
            if (needed > available)
            {
                var fitted = (available - textHeight) / (lines.Count - 1);
                if (fitted >= lineHeight * 0.85f)
                {
                    lineHeight = fitted;
                }
            }
        }

        var blockHeight = (lines.Count - 1) * lineHeight + textHeight;
        var top = style.Vertical switch
        {
            VerticalAlignment.Top => rect.Top + TextPadding,
            VerticalAlignment.Center => rect.Top + (rect.Height - blockHeight) / 2f,
            _ => rect.Bottom - blockHeight - 1f,
        };

        if (top < rect.Top)
        {
            top = rect.Top;
        }

        var isNumeric = cell.Kind == CellValueKind.Number;
        var alignment = style.Horizontal switch
        {
            HorizontalAlignment.General => isNumeric ? HorizontalAlignment.Right : HorizontalAlignment.Left,
            _ => style.Horizontal,
        };

        var textAlign = alignment switch
        {
            HorizontalAlignment.Center or HorizontalAlignment.CenterContinuous => SKTextAlign.Center,
            HorizontalAlignment.Right => SKTextAlign.Right,
            _ => SKTextAlign.Left,
        };
        paint.TextAlign = textAlign;

        var x = textAlign switch
        {
            SKTextAlign.Center => rect.MidX,
            SKTextAlign.Right => rect.Right - TextPadding,
            _ => rect.Left + TextPadding + indent,
        };

        canvas.Save();
        canvas.ClipRect(rect);

        for (var i = 0; i < lines.Count; i++)
        {
            var line = lines[i];
            var baseline = top - fontMetrics.Ascent + i * lineHeight;
            if (baseline > rect.Bottom + lineHeight)
            {
                break;
            }

            if (line.Length > 0)
            {
                canvas.DrawText(line, x, baseline, paint);

                if (style.Underline || style.Strike)
                {
                    var lineWidth = paint.MeasureText(line);
                    var left = textAlign switch
                    {
                        SKTextAlign.Center => x - lineWidth / 2f,
                        SKTextAlign.Right => x - lineWidth,
                        _ => x,
                    };

                    using var decorationPaint = new SKPaint
                    {
                        Style = SKPaintStyle.Stroke,
                        StrokeWidth = Math.Max(1f, fontPx / 14f),
                        Color = paint.Color,
                        IsAntialias = true,
                    };

                    if (style.Underline)
                    {
                        canvas.DrawLine(left, baseline + 2f, left + lineWidth, baseline + 2f, decorationPaint);
                    }

                    if (style.Strike)
                    {
                        var y = baseline - fontPx * 0.3f;
                        canvas.DrawLine(left, y, left + lineWidth, y, decorationPaint);
                    }
                }
            }
        }

        canvas.Restore();
    }

    private SKPaint CreateTextPaint(CellStyle style, float fontPx, string text)
    {
        return new SKPaint
        {
            IsAntialias = true,
            SubpixelText = true,
            Color = ToSkColor(style.FontColor),
            Typeface = GetTypeface(style, text),
            TextSize = fontPx,
            Style = SKPaintStyle.Fill,
        };
    }

    private static SKTypeface GetTypeface(CellStyle style, string text)
    {
        var typeface = FontResolver.Resolve(style.FontName, style.Bold, style.Italic);
        if (ContainsCjk(text) && !FontResolver.SupportsCjk(typeface, text))
        {
            typeface = FontResolver.ResolveCjk(style.Bold, style.Italic);
        }

        return typeface;
    }

    private SKPaint GetFillPaint(uint color)
    {
        if (_fillPaints.TryGetValue(color, out var paint))
        {
            return paint;
        }

        paint = new SKPaint
        {
            Style = SKPaintStyle.Fill,
            Color = ToSkColor(color),
            IsAntialias = false,
        };
        _fillPaints[color] = paint;
        return paint;
    }

    private static void DrawBorders(
        SKCanvas canvas,
        SKPaint paint,
        SKRect rect,
        BorderSide top,
        BorderSide right,
        BorderSide bottom,
        BorderSide left)
    {
        DrawBorderSide(canvas, paint, top, new SKPoint(rect.Left, rect.Top), new SKPoint(rect.Right, rect.Top));
        DrawBorderSide(canvas, paint, right, new SKPoint(rect.Right, rect.Top), new SKPoint(rect.Right, rect.Bottom));
        DrawBorderSide(canvas, paint, bottom, new SKPoint(rect.Left, rect.Bottom), new SKPoint(rect.Right, rect.Bottom));
        DrawBorderSide(canvas, paint, left, new SKPoint(rect.Left, rect.Top), new SKPoint(rect.Left, rect.Bottom));
    }

    private static void DrawBorderSide(SKCanvas canvas, SKPaint paint, BorderSide side, SKPoint start, SKPoint end)
    {
        if (!side.Visible)
        {
            return;
        }

        paint.Reset();
        paint.IsAntialias = true;
        paint.Style = SKPaintStyle.Stroke;
        paint.Color = ToSkColor(side.Color);
        paint.StrokeCap = SKStrokeCap.Butt;

        var (width, dash) = side.StyleIndex switch
        {
            1 => (1f, null),
            2 => (2f, null),
            3 => (1f, new[] { 4f, 3f }),
            4 => (1f, new[] { 1f, 2f }),
            5 => (3f, null),
            7 => (0.7f, null),
            8 => (2f, new[] { 5f, 3f }),
            9 => (1f, new[] { 4f, 2f, 1f, 2f }),
            10 => (2f, new[] { 5f, 2f, 1f, 2f }),
            11 => (1f, new[] { 4f, 2f, 1f, 2f, 1f, 2f }),
            12 => (2f, new[] { 5f, 2f, 1f, 2f, 1f, 2f }),
            13 => (1f, new[] { 6f, 2f, 1f, 2f }),
            _ => (1f, null),
        };

        paint.StrokeWidth = width;

        if (side.StyleIndex == 6)
        {
            // 双线：在两侧各画一条细线
            var dx = 0f;
            var dy = 0f;
            if (Math.Abs(end.Y - start.Y) < 0.01f)
            {
                dy = 1.2f;
            }
            else
            {
                dx = 1.2f;
            }

            paint.StrokeWidth = 1f;
            canvas.DrawLine(start.X - dx, start.Y - dy, end.X - dx, end.Y - dy, paint);
            canvas.DrawLine(start.X + dx, start.Y + dy, end.X + dx, end.Y + dy, paint);
            return;
        }

        if (dash != null)
        {
            using var effect = SKPathEffect.CreateDash(dash, 0);
            paint.PathEffect = effect;
            canvas.DrawLine(start, end, paint);
            paint.PathEffect = null;
            return;
        }

        canvas.DrawLine(start, end, paint);
    }

    private void BuildLayout()
    {
        if (_sheet.IsEmpty)
        {
            _contentWidth = 200;
            _contentHeight = 80;
            return;
        }

        for (var col = _sheet.MinCol; col <= _sheet.MaxCol; col++)
        {
            _colWidths[col] = GetColumnWidth(col);
        }

        _colX = new double[_sheet.MaxCol + 2];
        double x = 0;
        for (var col = _sheet.MinCol; col <= _sheet.MaxCol; col++)
        {
            _colX[col] = x;
            x += _colWidths[col];
        }

        _contentWidth = x;

        var cellsByRow = new Dictionary<int, List<CellData>>();
        var cellColsByRow = new Dictionary<int, List<int>>();
        foreach (var entry in _sheet.Cells)
        {
            if (!cellsByRow.TryGetValue(entry.Key.Row, out var cells))
            {
                cells = new List<CellData>();
                cellColsByRow[entry.Key.Row] = new List<int>();
                cellsByRow[entry.Key.Row] = cells;
            }

            cells.Add(entry.Value);
            cellColsByRow[entry.Key.Row].Add(entry.Key.Col);
        }

        _rowHeights.Clear();
        for (var row = _sheet.MinRow; row <= _sheet.MaxRow; row++)
        {
            var explicitHeight = _sheet.Rows.TryGetValue(row, out var rowInfo) ? rowInfo.HeightPt : null;
            var height = (explicitHeight ?? _sheet.DefaultRowHeightPt) * 4.0 / 3.0;

            if (explicitHeight == null && cellsByRow.TryGetValue(row, out var cells))
            {
                var cols = cellColsByRow[row];
                for (var i = 0; i < cells.Count; i++)
                {
                    var cell = cells[i];
                    var style = GetStyle(cell.StyleIndex);
                    if (!style.WrapText)
                    {
                        continue;
                    }

                    var key = (row, cols[i]);
                    if (_mergeAnchors.TryGetValue(key, out var merge) && merge.LastRow > merge.FirstRow)
                    {
                        continue;
                    }

                    var text = NumberFormatter.FormatValue(cell, style, _workbook);
                    if (string.IsNullOrEmpty(text))
                    {
                        continue;
                    }

                    var width = GetCellWidthSpan(row, cols[i]);
                    var fontPx = (float)(style.FontSizePt * 4.0 / 3.0);
                    if (fontPx <= 0 || width <= TextPadding * 2)
                    {
                        continue;
                    }

                    using var paint = CreateTextPaint(style, fontPx, text);
                    var lineCount = WrapText(text, (float)width - TextPadding * 2f, paint).Count;
                    var needed = lineCount * GetLineHeight(style) + TextPadding * 2f;
                    if (needed > height)
                    {
                        height = needed;
                    }
                }
            }

            if (_sheet.Rows.TryGetValue(row, out var info) && info.Hidden)
            {
                height = 0;
            }

            _rowHeights[row] = height;
        }

        _rowY = new double[_sheet.MaxRow + 2];
        double y = 0;
        for (var row = _sheet.MinRow; row <= _sheet.MaxRow; row++)
        {
            _rowY[row] = y;
            y += _rowHeights[row];
        }

        _contentHeight = y;
    }

    private double GetColumnWidth(int col)
    {
        if (_sheet.Cols.TryGetValue(col, out var info))
        {
            if (info.Hidden)
            {
                return 0;
            }

            if (info.WidthChars is { } width)
            {
                return ConvertCustomColumnWidth(width);
            }
        }

        return ConvertDefaultColumnWidth(_sheet.DefaultColWidthChars);
    }

    /// <summary>
    /// xlsx 中 &lt;col width&gt; 的单位是“默认字体最大数字宽度的倍数”，
    /// 因此像素宽度 = width × 最大数字宽度。若仍按固定 7px/字符换算，
    /// 默认字体为“等线”等宽字符字体时列会变窄，导致换行后文字显示不全。
    /// </summary>
    private double ConvertCustomColumnWidth(double chars)
    {
        if (chars <= 0)
        {
            return 0;
        }

        if (chars < 1)
        {
            return Math.Truncate(chars * 12 + 0.5);
        }

        return Math.Truncate(chars * _maxDigitWidth);
    }

    private double ConvertDefaultColumnWidth(double chars)
    {
        if (chars <= 0)
        {
            return 0;
        }

        if (chars < 1)
        {
            return Math.Truncate(chars * 12 + 0.5);
        }

        // 默认列宽包含单元格左右的内边距（约 5 像素）
        return Math.Truncate(chars * _maxDigitWidth + 5);
    }

    private static int MeasureMaxDigitWidth(Workbook workbook)
    {
        var size = Math.Max(1.0, workbook.DefaultFontSizePt);
        var typeface = FontResolver.Resolve(workbook.DefaultFontName, false, false);
        using var paint = new SKPaint
        {
            Typeface = typeface,
            TextSize = (float)(size * 96.0 / 72.0),
            IsAntialias = false,
        };

        return Math.Max(1, (int)Math.Round(paint.MeasureText("0")));
    }

    private double GetCellWidthSpan(int row, int col)
    {
        if (_mergeAnchors.TryGetValue((row, col), out var merge))
        {
            double width = 0;
            for (var c = merge.FirstCol; c <= merge.LastCol; c++)
            {
                width += _colWidths.TryGetValue(c, out var w) ? w : 0;
            }

            return width;
        }

        return _colWidths.TryGetValue(col, out var value) ? value : 0;
    }

    private CellStyle GetStyle(int index) =>
        index >= 0 && index < _workbook.Styles.Count ? _workbook.Styles[index] : _workbook.Styles[0];

    private CellStyle? GetStyleAt(int row, int col) =>
        _sheet.Cells.TryGetValue((row, col), out var cell) ? GetStyle(cell.StyleIndex) : null;

    private float GetLineHeight(CellStyle style)
    {
        // Excel 的换行行距与工作簿默认行高、默认字号相关：
        // 11pt 等线（行高 14pt）约 18.7px，11pt Calibri（行高 15pt）约 20px。
        var baseRowHeight = _sheet.DefaultRowHeightPt > 0
            ? _sheet.DefaultRowHeightPt
            : Math.Max(1.0, _workbook.DefaultRowHeightPt);
        var defaultFontSize = Math.Max(1.0, _workbook.DefaultFontSizePt);
        var fontSize = Math.Max(0.5, style.FontSizePt);
        var lineHeightPx = baseRowHeight * (fontSize / defaultFontSize) * 96.0 / 72.0;
        var minPx = fontSize * 4.0 / 3.0 * 1.1;
        return (float)Math.Max(Math.Max(10.0, minPx), lineHeightPx);
    }

    private static List<string> SplitLines(string text)
    {
        return text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n').ToList();
    }

    private static List<string> WrapText(string text, float maxWidth, SKPaint paint)
    {
        var result = new List<string>();
        foreach (var rawLine in SplitLines(text))
        {
            if (rawLine.Length == 0)
            {
                result.Add("");
                continue;
            }

            var start = 0;
            while (start < rawLine.Length)
            {
                var end = start;
                var lastBreak = -1;
                var width = 0f;

                while (end < rawLine.Length)
                {
                    var charWidth = paint.MeasureText(rawLine[end].ToString());
                    if (width + charWidth > maxWidth && end > start)
                    {
                        break;
                    }

                    width += charWidth;
                    if (rawLine[end] == ' ')
                    {
                        lastBreak = end;
                    }

                    end++;
                }

                if (end == start)
                {
                    end = start + 1;
                }

                if (end < rawLine.Length && lastBreak > start)
                {
                    end = lastBreak;
                }

                result.Add(rawLine[start..end].TrimEnd());
                start = end;
                while (start < rawLine.Length && rawLine[start] == ' ')
                {
                    start++;
                }
            }
        }

        return result;
    }

    private static bool ContainsCjk(string text)
    {
        foreach (var ch in text)
        {
            if (ch >= 0x2E80)
            {
                return true;
            }
        }

        return false;
    }

    private static SKColor ToSkColor(uint color) => new(
        (byte)((color >> 16) & 0xFF),
        (byte)((color >> 8) & 0xFF),
        (byte)(color & 0xFF),
        (byte)((color >> 24) & 0xFF));

    public void Dispose()
    {
        foreach (var paint in _fillPaints.Values)
        {
            paint.Dispose();
        }

        _fillPaints.Clear();
    }
}
