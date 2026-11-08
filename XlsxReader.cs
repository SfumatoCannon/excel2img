using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Excel2Img;

internal static class XlsxReader
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PR = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly uint[] IndexedPalette =
    {
        0xFF000000, 0xFFFFFFFF, 0xFFFF0000, 0xFF00FF00, 0xFF0000FF, 0xFFFFFF00, 0xFFFF00FF, 0xFF00FFFF,
        0xFF000000, 0xFFFFFFFF, 0xFFFF0000, 0xFF00FF00, 0xFF0000FF, 0xFFFFFF00, 0xFFFF00FF, 0xFF00FFFF,
        0xFF800000, 0xFF008000, 0xFF000080, 0xFF808000, 0xFF800080, 0xFF008080, 0xFFC0C0C0, 0xFF808080,
        0xFF9999FF, 0xFF993366, 0xFFFFFFCC, 0xFFCCFFFF, 0xFF660066, 0xFFFF8080, 0xFF0066CC, 0xFFCCCCFF,
        0xFF000080, 0xFFFF00FF, 0xFFFFFF00, 0xFF00FFFF, 0xFF800080, 0xFF800000, 0xFF008080, 0xFF0000FF,
        0xFF00CCFF, 0xFFCCFFFF, 0xFFCCFFCC, 0xFFFFFF99, 0xFF99CCFF, 0xFFFF99CC, 0xFFCC99FF, 0xFFFFCC99,
        0xFF3366FF, 0xFF33CCCC, 0xFF99CC00, 0xFFFFCC00, 0xFFFF9900, 0xFFFF6600, 0xFF666699, 0xFF969696,
        0xFF003366, 0xFF339966, 0xFF003300, 0xFF333300, 0xFF993300, 0xFF993366, 0xFF333399, 0xFF333333,
        0xFF000000, 0xFFFFFFFF,
    };

    public static (Workbook Workbook, SheetInfo Sheet) Load(string filePath, string? sheetSelector)
    {
        if (!File.Exists(filePath))
        {
            throw new CliException(Localization.F(
                "Input file not found: {0}",
                "找不到输入文件: {0}",
                filePath));
        }

        if (!string.Equals(Path.GetExtension(filePath), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            throw new CliException(Localization.T(
                "Only .xlsx Excel files are supported.",
                "仅支持 .xlsx 格式的 Excel 文件。"));
        }

        try
        {
            using var zip = ZipFile.OpenRead(filePath);
            var workbook = new Workbook { FilePath = filePath };

            var workbookEntry = FindEntry(zip, "xl/workbook.xml")
                ?? throw new CliException(Localization.T(
                    "Cannot parse the file: xl/workbook.xml is missing; it may not be a valid .xlsx file.",
                    "无法解析文件：缺少 xl/workbook.xml，可能不是有效的 .xlsx 文件。"));

            var workbookDoc = LoadXml(workbookEntry);
            var workbookPr = workbookDoc.Root?.Element(M + "workbookPr");
            workbook.Date1904 = IsTrue(workbookPr?.Attribute("date1904")?.Value);

            var relationships = ReadRelationships(zip);
            var sheetsElement = workbookDoc.Root?.Element(M + "sheets");
            if (sheetsElement != null)
            {
                foreach (var sheetElement in sheetsElement.Elements(M + "sheet"))
                {
                    var name = sheetElement.Attribute("name")?.Value ?? "Sheet";
                    var relId = sheetElement.Attribute(R + "id")?.Value;
                    if (relId == null || !relationships.TryGetValue(relId, out var target))
                    {
                        continue;
                    }

                    var state = sheetElement.Attribute("state")?.Value ?? "visible";
                    workbook.Sheets.Add(new SheetMeta
                    {
                        Name = name,
                        Target = NormalizeTarget(target),
                        Hidden = !string.Equals(state, "visible", StringComparison.OrdinalIgnoreCase),
                    });
                }
            }

            if (workbook.Sheets.Count == 0)
            {
                throw new CliException(Localization.T(
                    "The workbook contains no worksheets.",
                    "工作簿中没有可用的工作表。"));
            }

            ReadSharedStrings(zip, workbook);
            ReadStyles(zip, workbook);
            ReadTheme(zip, workbook);

            var meta = ResolveSheet(workbook, sheetSelector);
            var sheet = ReadSheet(zip, workbook, meta);
            return (workbook, sheet);
        }
        catch (InvalidDataException)
        {
            throw new CliException(Localization.T(
                "Cannot open the file; it may be corrupted or not a valid .xlsx file.",
                "无法打开文件，文件可能已损坏或不是有效的 .xlsx 文件。"));
        }
        catch (IOException ex)
        {
            throw new CliException(Localization.F(
                "Failed to read the file: {0}",
                "读取文件失败: {0}",
                ex.Message));
        }
    }

    private static SheetMeta ResolveSheet(Workbook workbook, string? selector)
    {
        if (string.IsNullOrWhiteSpace(selector))
        {
            return workbook.Sheets.FirstOrDefault(s => !s.Hidden) ?? workbook.Sheets[0];
        }

        if (int.TryParse(selector, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index))
        {
            if (index < 1 || index > workbook.Sheets.Count)
            {
                throw new CliException(Localization.F(
                    "Sheet index out of range: {0} (the workbook has {1} worksheets).",
                    "工作表序号超出范围：{0}（共 {1} 个工作表）。",
                    index,
                    workbook.Sheets.Count));
            }

            return workbook.Sheets[index - 1];
        }

        var matched = workbook.Sheets.FirstOrDefault(s =>
            string.Equals(s.Name, selector, StringComparison.OrdinalIgnoreCase));
        if (matched == null)
        {
            var separator = Localization.T(", ", "、");
            var names = string.Join(separator, workbook.Sheets.Select(s => s.Name));
            throw new CliException(Localization.F(
                "Worksheet \"{0}\" was not found. Available worksheets: {1}",
                "找不到名为 “{0}” 的工作表。可用工作表：{1}",
                selector,
                names));
        }

        return matched;
    }

    private static void ReadSharedStrings(ZipArchive zip, Workbook workbook)
    {
        var entry = FindEntry(zip, "xl/sharedStrings.xml");
        if (entry == null)
        {
            return;
        }

        var doc = LoadXml(entry);
        foreach (var si in doc.Root?.Elements(M + "si") ?? Enumerable.Empty<XElement>())
        {
            var builder = new StringBuilder();
            var plain = si.Element(M + "t");
            if (plain != null)
            {
                builder.Append(plain.Value);
            }

            foreach (var run in si.Elements(M + "r"))
            {
                var text = run.Element(M + "t");
                if (text != null)
                {
                    builder.Append(text.Value);
                }
            }

            workbook.SharedStrings.Add(builder.ToString());
        }
    }

    private static void ReadStyles(ZipArchive zip, Workbook workbook)
    {
        var entry = FindEntry(zip, "xl/styles.xml");
        if (entry == null)
        {
            workbook.Styles.Add(new CellStyle());
            return;
        }

        var root = LoadXml(entry).Root;
        if (root == null)
        {
            workbook.Styles.Add(new CellStyle());
            return;
        }

        var customFormats = new Dictionary<int, string>();
        var numFmtsElement = root.Element(M + "numFmts");
        if (numFmtsElement != null)
        {
            foreach (var numFmt in numFmtsElement.Elements(M + "numFmt"))
            {
                if (int.TryParse(numFmt.Attribute("numFmtId")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id))
                {
                    customFormats[id] = numFmt.Attribute("formatCode")?.Value ?? "";
                }
            }
        }

        var fonts = new List<FontInfo>();
        var fontsElement = root.Element(M + "fonts");
        if (fontsElement != null)
        {
            foreach (var font in fontsElement.Elements(M + "font"))
            {
                fonts.Add(new FontInfo
                {
                    Name = font.Element(M + "name")?.Attribute("val")?.Value ?? "Calibri",
                    Size = ParseDouble(font.Element(M + "sz")?.Attribute("val")?.Value, 11),
                    Bold = HasToggle(font.Element(M + "b")),
                    Italic = HasToggle(font.Element(M + "i")),
                    Underline = HasUnderline(font.Element(M + "u")),
                    Strike = HasToggle(font.Element(M + "strike")),
                    Color = ReadColor(font.Element(M + "color"), workbook.Theme),
                });
            }
        }

        if (fonts.Count == 0)
        {
            fonts.Add(new FontInfo());
        }

        workbook.DefaultFontSizePt = fonts[0].Size;
        workbook.DefaultFontName = fonts[0].Name;

        var fills = new List<uint?>();
        var fillsElement = root.Element(M + "fills");
        if (fillsElement != null)
        {
            foreach (var fill in fillsElement.Elements(M + "fill"))
            {
                var pattern = fill.Element(M + "patternFill");
                var patternType = pattern?.Attribute("patternType")?.Value;
                uint? color = null;
                if (pattern != null && !string.Equals(patternType, "none", StringComparison.OrdinalIgnoreCase))
                {
                    color = ReadColor(pattern.Element(M + "fgColor"), workbook.Theme)
                        ?? ReadColor(pattern.Element(M + "bgColor"), workbook.Theme);
                }

                fills.Add(color);
            }
        }

        var borders = new List<BorderInfo>();
        var bordersElement = root.Element(M + "borders");
        if (bordersElement != null)
        {
            foreach (var border in bordersElement.Elements(M + "border"))
            {
                borders.Add(new BorderInfo
                {
                    Left = ReadBorderSide(border.Element(M + "left"), workbook.Theme),
                    Right = ReadBorderSide(border.Element(M + "right"), workbook.Theme),
                    Top = ReadBorderSide(border.Element(M + "top"), workbook.Theme),
                    Bottom = ReadBorderSide(border.Element(M + "bottom"), workbook.Theme),
                });
            }
        }

        if (borders.Count == 0)
        {
            borders.Add(new BorderInfo());
        }

        var cellXfsElement = root.Element(M + "cellXfs");
        if (cellXfsElement == null)
        {
            workbook.Styles.Add(new CellStyle());
            return;
        }

        foreach (var xf in cellXfsElement.Elements(M + "xf"))
        {
            var style = new CellStyle
            {
                NumFmtId = ParseInt(xf.Attribute("numFmtId")?.Value, 0),
                FormatCode = null,
            };

            if (customFormats.TryGetValue(style.NumFmtId, out var code))
            {
                style.FormatCode = code;
            }

            var fontId = ParseInt(xf.Attribute("fontId")?.Value, 0);
            if (fontId >= 0 && fontId < fonts.Count)
            {
                var font = fonts[fontId];
                style.FontName = font.Name;
                style.FontSizePt = font.Size;
                style.Bold = font.Bold;
                style.Italic = font.Italic;
                style.Underline = font.Underline;
                style.Strike = font.Strike;
                style.FontColor = font.Color ?? 0xFF000000;
            }

            var fillId = ParseInt(xf.Attribute("fillId")?.Value, 0);
            if (fillId >= 0 && fillId < fills.Count)
            {
                style.FillColor = fills[fillId];
            }

            var borderId = ParseInt(xf.Attribute("borderId")?.Value, 0);
            if (borderId >= 0 && borderId < borders.Count)
            {
                var border = borders[borderId];
                style.Left = border.Left;
                style.Right = border.Right;
                style.Top = border.Top;
                style.Bottom = border.Bottom;
            }

            var alignment = xf.Element(M + "alignment");
            if (alignment != null)
            {
                style.Horizontal = ParseHorizontal(alignment.Attribute("horizontal")?.Value);
                style.Vertical = ParseVertical(alignment.Attribute("vertical")?.Value);
                style.WrapText = IsTrue(alignment.Attribute("wrapText")?.Value);
                style.Indent = ParseInt(alignment.Attribute("indent")?.Value, 0);
            }

            workbook.Styles.Add(style);
        }

        if (workbook.Styles.Count == 0)
        {
            workbook.Styles.Add(new CellStyle());
        }
    }

    private static void ReadTheme(ZipArchive zip, Workbook workbook)
    {
        var entry = FindEntry(zip, "xl/theme/theme1.xml")
            ?? zip.Entries.FirstOrDefault(e => e.FullName.StartsWith("xl/theme/", StringComparison.OrdinalIgnoreCase));
        if (entry == null)
        {
            return;
        }

        var doc = LoadXml(entry);
        var scheme = doc.Descendants(A + "clrScheme").FirstOrDefault();
        if (scheme == null)
        {
            return;
        }

        var indices = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["lt1"] = 0,
            ["dk1"] = 1,
            ["lt2"] = 2,
            ["dk2"] = 3,
            ["accent1"] = 4,
            ["accent2"] = 5,
            ["accent3"] = 6,
            ["accent4"] = 7,
            ["accent5"] = 8,
            ["accent6"] = 9,
            ["hlink"] = 10,
            ["folHlink"] = 11,
        };

        foreach (var child in scheme.Elements())
        {
            if (!indices.TryGetValue(child.Name.LocalName, out var index))
            {
                continue;
            }

            var srgb = child.Element(A + "srgbClr")?.Attribute("val")?.Value
                ?? child.Element(A + "sysClr")?.Attribute("lastClr")?.Value;
            if (!string.IsNullOrEmpty(srgb) && TryParseHexColor(srgb, out var color))
            {
                workbook.Theme.Set(index, color);
            }
        }
    }

    private static SheetInfo ReadSheet(ZipArchive zip, Workbook workbook, SheetMeta meta)
    {
        var entry = FindEntry(zip, meta.Target)
            ?? throw new CliException(Localization.F(
                "Cannot find data for worksheet \"{0}\".",
                "找不到工作表 “{0}” 的数据。",
                meta.Name));

        var root = LoadXml(entry).Root
            ?? throw new CliException(Localization.F(
                "Worksheet \"{0}\" is empty.",
                "工作表 “{0}” 的内容为空。",
                meta.Name));

        var sheet = new SheetInfo
        {
            Name = meta.Name,
            Hidden = meta.Hidden,
            DefaultRowHeightPt = workbook.DefaultRowHeightPt,
            DefaultColWidthChars = workbook.DefaultColWidthChars,
        };

        var sheetFormat = root.Element(M + "sheetFormatPr");
        if (sheetFormat != null)
        {
            sheet.DefaultRowHeightPt = ParseDouble(sheetFormat.Attribute("defaultRowHeight")?.Value, workbook.DefaultRowHeightPt);
            sheet.DefaultColWidthChars = ParseDouble(sheetFormat.Attribute("defaultColWidth")?.Value, workbook.DefaultColWidthChars);
        }

        var sheetView = root.Element(M + "sheetViews")?.Element(M + "sheetView");
        if (sheetView?.Attribute("showGridLines") is { } gridLines)
        {
            sheet.ShowGridLines = !string.Equals(gridLines.Value, "0", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(gridLines.Value, "false", StringComparison.OrdinalIgnoreCase);
        }

        var colsElement = root.Element(M + "cols");
        if (colsElement != null)
        {
            foreach (var col in colsElement.Elements(M + "col"))
            {
                var min = ParseInt(col.Attribute("min")?.Value, 0);
                var max = ParseInt(col.Attribute("max")?.Value, 0);
                if (min <= 0 || max < min)
                {
                    continue;
                }

                var info = new ColInfo
                {
                    WidthChars = col.Attribute("width") is { } width ? ParseDouble(width.Value, sheet.DefaultColWidthChars) : null,
                    Hidden = IsTrue(col.Attribute("hidden")?.Value),
                };

                for (var i = min; i <= Math.Min(max, 16384); i++)
                {
                    sheet.Cols[i] = info;
                }
            }
        }

        var sheetData = root.Element(M + "sheetData");
        if (sheetData != null)
        {
            var rowIndex = 0;
            foreach (var rowElement in sheetData.Elements(M + "row"))
            {
                rowIndex = ParseInt(rowElement.Attribute("r")?.Value, rowIndex + 1);
                var rowInfo = new RowInfo
                {
                    HeightPt = rowElement.Attribute("ht") is { } ht ? ParseDouble(ht.Value, sheet.DefaultRowHeightPt) : null,
                    Hidden = IsTrue(rowElement.Attribute("hidden")?.Value),
                };
                sheet.Rows[rowIndex] = rowInfo;

                var colIndex = 0;
                foreach (var cellElement in rowElement.Elements(M + "c"))
                {
                    var reference = cellElement.Attribute("r")?.Value;
                    int row, col;
                    if (!string.IsNullOrEmpty(reference))
                    {
                        ParseCellReference(reference, out row, out col);
                    }
                    else
                    {
                        row = rowIndex;
                        col = colIndex + 1;
                    }

                    colIndex = col;

                    var cell = new CellData
                    {
                        StyleIndex = ParseInt(cellElement.Attribute("s")?.Value, 0),
                    };

                    var type = cellElement.Attribute("t")?.Value;
                    var valueElement = cellElement.Element(M + "v");
                    var value = valueElement?.Value;

                    switch (type)
                    {
                        case "s":
                            cell.Kind = CellValueKind.SharedString;
                            cell.SharedStringIndex = ParseInt(value, -1);
                            break;
                        case "inlineStr":
                            cell.Kind = CellValueKind.InlineString;
                            cell.RawValue = ReadInlineString(cellElement.Element(M + "is"));
                            break;
                        case "str":
                            cell.Kind = CellValueKind.String;
                            cell.RawValue = value ?? "";
                            break;
                        case "b":
                            cell.Kind = CellValueKind.Bool;
                            cell.RawValue = value ?? "";
                            break;
                        case "e":
                            cell.Kind = CellValueKind.Error;
                            cell.RawValue = value ?? "";
                            break;
                        case "d":
                            cell.Kind = CellValueKind.String;
                            cell.RawValue = value ?? "";
                            break;
                        default:
                            if (string.IsNullOrEmpty(value))
                            {
                                cell.Kind = CellValueKind.Empty;
                            }
                            else
                            {
                                cell.Kind = CellValueKind.Number;
                                cell.RawValue = value;
                            }

                            break;
                    }

                    if (cell.Kind == CellValueKind.Empty && cell.StyleIndex == 0)
                    {
                        // 没有内容也没有样式的空单元格无需渲染
                        continue;
                    }

                    sheet.Cells[(row, col)] = cell;
                    UpdateBounds(sheet, row, col);
                }
            }
        }

        var mergeCellsElement = root.Element(M + "mergeCells");
        if (mergeCellsElement != null)
        {
            foreach (var mergeElement in mergeCellsElement.Elements(M + "mergeCell"))
            {
                var reference = mergeElement.Attribute("ref")?.Value;
                if (string.IsNullOrEmpty(reference))
                {
                    continue;
                }

                var parts = reference.Split(':');
                if (parts.Length != 2)
                {
                    continue;
                }

                ParseCellReference(parts[0], out var firstRow, out var firstCol);
                ParseCellReference(parts[1], out var lastRow, out var lastCol);
                if (lastRow < firstRow || lastCol < firstCol)
                {
                    continue;
                }

                sheet.Merged.Add(new MergedRange(firstRow, firstCol, lastRow, lastCol));
                UpdateBounds(sheet, firstRow, firstCol);
                UpdateBounds(sheet, lastRow, lastCol);
            }
        }

        ReadConditionalFormatting(root, workbook, sheet);
        ApplyColorScales(sheet);

        return sheet;
    }

    private static void ReadConditionalFormatting(XElement root, Workbook workbook, SheetInfo sheet)
    {
        foreach (var formatting in root.Elements(M + "conditionalFormatting"))
        {
            var sqref = formatting.Attribute("sqref")?.Value;
            if (string.IsNullOrWhiteSpace(sqref))
            {
                continue;
            }

            var ranges = ParseRanges(sqref);
            if (ranges.Count == 0)
            {
                continue;
            }

            foreach (var rule in formatting.Elements(M + "cfRule"))
            {
                if (!string.Equals(rule.Attribute("type")?.Value, "colorScale", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var scale = rule.Element(M + "colorScale");
                if (scale == null)
                {
                    continue;
                }

                var colors = scale.Elements(M + "color")
                    .Select(color => ReadColor(color, workbook.Theme) ?? 0xFF000000)
                    .ToList();
                if (colors.Count < 2)
                {
                    continue;
                }

                var cfvos = scale.Elements(M + "cfvo").ToList();
                var colorScale = new ColorScaleRule
                {
                    Priority = ParseInt(rule.Attribute("priority")?.Value, int.MaxValue),
                };
                colorScale.Ranges.AddRange(ranges);

                for (var i = 0; i < colors.Count; i++)
                {
                    var stop = new ColorScaleStop
                    {
                        Color = colors[i],
                        Type = i < cfvos.Count ? (cfvos[i].Attribute("type")?.Value ?? "") : "",
                    };

                    if (i < cfvos.Count
                        && double.TryParse(cfvos[i].Attribute("val")?.Value, NumberStyles.Float,
                            CultureInfo.InvariantCulture, out var value))
                    {
                        stop.Value = value;
                    }

                    if (string.IsNullOrEmpty(stop.Type))
                    {
                        stop.Type = i == 0 ? "min" : i == colors.Count - 1 ? "max" : "percentile";
                    }

                    colorScale.Stops.Add(stop);
                }

                sheet.ColorScales.Add(colorScale);
            }
        }
    }

    private static List<MergedRange> ParseRanges(string sqref)
    {
        var ranges = new List<MergedRange>();
        foreach (var token in sqref.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = token.Split(':');
            if (parts.Length == 1)
            {
                ParseCellReference(parts[0], out var row, out var col);
                if (row > 0 && col > 0)
                {
                    ranges.Add(new MergedRange(row, col, row, col));
                }

                continue;
            }

            if (parts.Length != 2)
            {
                continue;
            }

            ParseCellReference(parts[0], out var firstRow, out var firstCol);
            ParseCellReference(parts[1], out var lastRow, out var lastCol);
            if (firstRow <= 0 || firstCol <= 0 || lastRow <= 0 || lastCol <= 0)
            {
                continue;
            }

            ranges.Add(new MergedRange(
                Math.Min(firstRow, lastRow),
                Math.Min(firstCol, lastCol),
                Math.Max(firstRow, lastRow),
                Math.Max(firstCol, lastCol)));
        }

        return ranges;
    }

    private static void ApplyColorScales(SheetInfo sheet)
    {
        if (sheet.ColorScales.Count == 0)
        {
            return;
        }

        foreach (var rule in sheet.ColorScales.OrderBy(r => r.Priority))
        {
            var values = new List<double>();
            foreach (var entry in sheet.Cells)
            {
                if (entry.Value.Kind != CellValueKind.Number || !IsInRanges(rule, entry.Key.Row, entry.Key.Col))
                {
                    continue;
                }

                if (double.TryParse(entry.Value.RawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
                {
                    values.Add(value);
                }
            }

            if (values.Count == 0)
            {
                continue;
            }

            values.Sort();
            var min = values[0];
            var max = values[^1];
            var positions = new double[rule.Stops.Count];

            for (var i = 0; i < rule.Stops.Count; i++)
            {
                var stop = rule.Stops[i];
                positions[i] = stop.Type.ToLowerInvariant() switch
                {
                    "min" => min,
                    "max" => max,
                    "num" => stop.Value ?? (i == 0 ? min : max),
                    "percent" => stop.Value is { } percent ? min + (max - min) * percent / 100.0 : min,
                    "percentile" => stop.Value is { } percentile ? Percentile(values, percentile) : min,
                    _ => stop.Value ?? (i == 0 ? min : max),
                };
            }

            for (var i = 1; i < positions.Length; i++)
            {
                if (positions[i] < positions[i - 1])
                {
                    positions[i] = positions[i - 1];
                }
            }

            var colors = rule.Stops.Select(stop => stop.Color).ToArray();
            foreach (var entry in sheet.Cells)
            {
                var key = entry.Key;
                if (sheet.ConditionalFills.ContainsKey(key)
                    || entry.Value.Kind != CellValueKind.Number
                    || !IsInRanges(rule, key.Row, key.Col)
                    || !double.TryParse(entry.Value.RawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
                {
                    continue;
                }

                sheet.ConditionalFills[key] = InterpolateColor(positions, colors, value);
            }
        }
    }

    private static bool IsInRanges(ColorScaleRule rule, int row, int col)
    {
        foreach (var range in rule.Ranges)
        {
            if (range.Contains(row, col))
            {
                return true;
            }
        }

        return false;
    }

    private static double Percentile(List<double> sortedValues, double percentile)
    {
        if (sortedValues.Count == 0)
        {
            return 0;
        }

        if (sortedValues.Count == 1)
        {
            return sortedValues[0];
        }

        var rank = Math.Clamp(percentile, 0, 100) / 100.0 * (sortedValues.Count - 1);
        var lower = (int)Math.Floor(rank);
        var upper = Math.Min(lower + 1, sortedValues.Count - 1);
        var fraction = rank - lower;
        return sortedValues[lower] + (sortedValues[upper] - sortedValues[lower]) * fraction;
    }

    private static uint InterpolateColor(double[] positions, uint[] colors, double value)
    {
        if (positions.Length == 0 || colors.Length == 0)
        {
            return 0xFFFFFFFF;
        }

        if (value <= positions[0] || positions.Length == 1)
        {
            return colors[0];
        }

        if (value >= positions[^1])
        {
            return colors[^1];
        }

        for (var i = 0; i < positions.Length - 1 && i + 1 < colors.Length; i++)
        {
            if (value > positions[i + 1])
            {
                continue;
            }

            var span = positions[i + 1] - positions[i];
            var ratio = span <= 0 ? 0 : (value - positions[i]) / span;
            return LerpColor(colors[i], colors[i + 1], ratio);
        }

        return colors[^1];
    }

    private static uint LerpColor(uint from, uint to, double ratio)
    {
        byte Lerp(byte start, byte end) =>
            (byte)Math.Clamp(Math.Round(start + (end - start) * ratio), 0, 255);

        var alpha = Lerp((byte)((from >> 24) & 0xFF), (byte)((to >> 24) & 0xFF));
        var red = Lerp((byte)((from >> 16) & 0xFF), (byte)((to >> 16) & 0xFF));
        var green = Lerp((byte)((from >> 8) & 0xFF), (byte)((to >> 8) & 0xFF));
        var blue = Lerp((byte)(from & 0xFF), (byte)(to & 0xFF));
        return ((uint)alpha << 24) | ((uint)red << 16) | ((uint)green << 8) | blue;
    }

    private static void UpdateBounds(SheetInfo sheet, int row, int col)
    {
        if (row < sheet.MinRow) sheet.MinRow = row;
        if (col < sheet.MinCol) sheet.MinCol = col;
        if (row > sheet.MaxRow) sheet.MaxRow = row;
        if (col > sheet.MaxCol) sheet.MaxCol = col;
    }

    private static string ReadInlineString(XElement? inline)
    {
        if (inline == null)
        {
            return "";
        }

        var builder = new StringBuilder();
        var plain = inline.Element(M + "t");
        if (plain != null)
        {
            builder.Append(plain.Value);
        }

        foreach (var run in inline.Elements(M + "r"))
        {
            var text = run.Element(M + "t");
            if (text != null)
            {
                builder.Append(text.Value);
            }
        }

        return builder.ToString();
    }

    private static BorderSide ReadBorderSide(XElement? element, ThemeColors theme)
    {
        var side = new BorderSide();
        if (element == null)
        {
            return side;
        }

        var styleName = element.Attribute("style")?.Value;
        if (string.IsNullOrEmpty(styleName) || string.Equals(styleName, "none", StringComparison.OrdinalIgnoreCase))
        {
            return side;
        }

        side.Visible = true;
        side.StyleIndex = styleName.ToLowerInvariant() switch
        {
            "thin" => 1,
            "medium" => 2,
            "dashed" => 3,
            "dotted" => 4,
            "thick" => 5,
            "double" => 6,
            "hair" => 7,
            "mediumdashed" => 8,
            "dashdot" => 9,
            "mediumdashdot" => 10,
            "dashdotdot" => 11,
            "mediumdashdotdot" => 12,
            "slantdashdot" => 13,
            _ => 1,
        };
        side.Color = ReadColor(element.Element(M + "color"), theme) ?? 0xFF000000;
        return side;
    }

    private static uint? ReadColor(XElement? element, ThemeColors theme)
    {
        if (element == null)
        {
            return null;
        }

        uint color;
        var rgb = element.Attribute("rgb")?.Value;
        if (!string.IsNullOrEmpty(rgb))
        {
            if (!TryParseHexColor(rgb, out color))
            {
                return null;
            }
        }
        else if (element.Attribute("theme") is { } themeAttr
                 && int.TryParse(themeAttr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var themeIndex))
        {
            color = theme.Get(themeIndex);
        }
        else if (element.Attribute("indexed") is { } indexedAttr
                 && int.TryParse(indexedAttr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indexed))
        {
            color = indexed >= 0 && indexed < IndexedPalette.Length ? IndexedPalette[indexed] : 0xFF000000;
        }
        else
        {
            return null;
        }

        if (element.Attribute("tint") is { } tintAttr
            && double.TryParse(tintAttr.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var tint))
        {
            color = ThemeColors.ApplyTint(color, tint);
        }

        return color;
    }

    private static bool TryParseHexColor(string value, out uint color)
    {
        value = value.Trim();
        if (value.StartsWith("#", StringComparison.Ordinal))
        {
            value = value[1..];
        }

        if (value.Length == 6)
        {
            value = "FF" + value;
        }

        if (value.Length == 8 && uint.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out color))
        {
            return true;
        }

        color = 0xFF000000;
        return false;
    }

    private static HorizontalAlignment ParseHorizontal(string? value) => value switch
    {
        "left" => HorizontalAlignment.Left,
        "center" => HorizontalAlignment.Center,
        "right" => HorizontalAlignment.Right,
        "fill" => HorizontalAlignment.Fill,
        "justify" => HorizontalAlignment.Justify,
        "centerContinuous" => HorizontalAlignment.CenterContinuous,
        "distributed" => HorizontalAlignment.Distributed,
        _ => HorizontalAlignment.General,
    };

    private static VerticalAlignment ParseVertical(string? value) => value switch
    {
        "top" => VerticalAlignment.Top,
        "center" => VerticalAlignment.Center,
        "justify" => VerticalAlignment.Justify,
        "distributed" => VerticalAlignment.Distributed,
        _ => VerticalAlignment.Bottom,
    };

    private static bool HasToggle(XElement? element)
    {
        if (element == null)
        {
            return false;
        }

        var value = element.Attribute("val")?.Value;
        return value == null || IsTrue(value);
    }

    private static bool HasUnderline(XElement? element)
    {
        if (element == null)
        {
            return false;
        }

        var value = element.Attribute("val")?.Value;
        return value == null || (!string.Equals(value, "none", StringComparison.OrdinalIgnoreCase)
                                 && !string.Equals(value, "0", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsTrue(string? value) =>
        value != null && (value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase));

    private static int ParseInt(string? value, int fallback) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result) ? result : fallback;

    private static double ParseDouble(string? value, double fallback) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result) ? result : fallback;

    private static void ParseCellReference(string reference, out int row, out int col)
    {
        reference = reference.Replace("$", "");
        row = 0;
        col = 0;
        var index = 0;
        while (index < reference.Length && char.IsLetter(reference[index]))
        {
            col = col * 26 + (char.ToUpperInvariant(reference[index]) - 'A' + 1);
            index++;
        }

        int.TryParse(reference[index..], NumberStyles.Integer, CultureInfo.InvariantCulture, out row);
    }

    private static Dictionary<string, string> ReadRelationships(ZipArchive zip)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var entry = FindEntry(zip, "xl/_rels/workbook.xml.rels");
        if (entry == null)
        {
            return result;
        }

        var doc = LoadXml(entry);
        foreach (var relationship in doc.Root?.Elements(PR + "Relationship") ?? Enumerable.Empty<XElement>())
        {
            var id = relationship.Attribute("Id")?.Value;
            if (!string.IsNullOrEmpty(id))
            {
                result[id] = relationship.Attribute("Target")?.Value ?? "";
            }
        }

        return result;
    }

    private static string NormalizeTarget(string target)
    {
        target = target.Replace('\\', '/');
        if (target.StartsWith("/", StringComparison.Ordinal))
        {
            return target.TrimStart('/');
        }

        return "xl/" + target;
    }

    private static ZipArchiveEntry? FindEntry(ZipArchive zip, string name) =>
        zip.GetEntry(name)
        ?? zip.Entries.FirstOrDefault(e => string.Equals(e.FullName, name, StringComparison.OrdinalIgnoreCase));

    private static XDocument LoadXml(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.None);
    }

    private sealed class FontInfo
    {
        public string Name = "Calibri";
        public double Size = 11;
        public bool Bold;
        public bool Italic;
        public bool Underline;
        public bool Strike;
        public uint? Color;
    }

    private sealed class BorderInfo
    {
        public BorderSide Left = new();
        public BorderSide Right = new();
        public BorderSide Top = new();
        public BorderSide Bottom = new();
    }
}
