namespace Excel2Img;

internal enum CellValueKind
{
    Empty,
    Number,
    SharedString,
    InlineString,
    String,
    Bool,
    Error,
}

internal sealed class CellData
{
    public CellValueKind Kind;
    public string RawValue = "";
    public int StyleIndex;
    public int SharedStringIndex = -1;
}

internal sealed class RowInfo
{
    public double? HeightPt;
    public bool Hidden;
}

internal sealed class ColInfo
{
    public double? WidthChars;
    public bool Hidden;
}

internal readonly record struct MergedRange(int FirstRow, int FirstCol, int LastRow, int LastCol)
{
    public bool Contains(int row, int col) =>
        row >= FirstRow && row <= LastRow && col >= FirstCol && col <= LastCol;
}

internal sealed class BorderSide
{
    public bool Visible;
    public int StyleIndex;
    public uint Color = 0xFF000000;
}

internal enum HorizontalAlignment
{
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed,
}

internal enum VerticalAlignment
{
    Bottom,
    Center,
    Top,
    Justify,
    Distributed,
}

internal sealed class CellStyle
{
    public int NumFmtId;
    public string? FormatCode;

    public string FontName = "Calibri";
    public double FontSizePt = 11;
    public bool Bold;
    public bool Italic;
    public bool Underline;
    public bool Strike;
    public uint FontColor = 0xFF000000;

    public uint? FillColor;

    public HorizontalAlignment Horizontal = HorizontalAlignment.General;
    public VerticalAlignment Vertical = VerticalAlignment.Bottom;
    public bool WrapText;
    public int Indent;

    public BorderSide Top = new();
    public BorderSide Right = new();
    public BorderSide Bottom = new();
    public BorderSide Left = new();
}

internal sealed class ThemeColors
{
    // 索引与 xlsx 中 theme 属性一致：0=lt1, 1=dk1, 2=lt2, 3=dk2, 4-9=accent1-6, 10=hlink, 11=folHlink
    private readonly uint[] _colors =
    {
        0xFFFFFFFF, 0xFF000000, 0xFFE7E6E6, 0xFF44546A,
        0xFF4472C4, 0xFFED7D31, 0xFFA5A5A5, 0xFFFFC000,
        0xFF5B9BD5, 0xFF70AD47, 0xFF0563C1, 0xFF954F72,
    };

    public uint Get(int index) =>
        index >= 0 && index < _colors.Length ? _colors[index] : 0xFF000000;

    public void Set(int index, uint color)
    {
        if (index >= 0 && index < _colors.Length)
        {
            _colors[index] = color;
        }
    }

    // 依据 Excel 的 tint 规则做近似调整
    public static uint ApplyTint(uint color, double tint)
    {
        if (tint == 0) return color;

        byte Adjust(byte channel)
        {
            double value = tint < 0
                ? channel * (1.0 + tint)
                : channel + (255.0 - channel) * tint;
            return (byte)Math.Clamp(Math.Round(value), 0, 255);
        }

        uint a = (color >> 24) & 0xFF;
        uint r = (color >> 16) & 0xFF;
        uint g = (color >> 8) & 0xFF;
        uint b = color & 0xFF;
        return (a << 24) | ((uint)Adjust((byte)r) << 16) | ((uint)Adjust((byte)g) << 8) | Adjust((byte)b);
    }
}

internal sealed class SheetMeta
{
    public string Name = "";
    public string Target = "";
    public bool Hidden;
}

internal sealed class Workbook
{
    public string FilePath = "";
    public readonly List<SheetMeta> Sheets = new();
    public readonly List<string> SharedStrings = new();
    public readonly List<CellStyle> Styles = new();
    public readonly ThemeColors Theme = new();
    public bool Date1904;
    public double DefaultRowHeightPt = 15;
    public double DefaultColWidthChars = 8.43;
    public double DefaultFontSizePt = 11;
    public string DefaultFontName = "Calibri";
}

internal sealed class SheetInfo
{
    public string Name = "";
    public bool Hidden;
    public bool ShowGridLines = true;
    public readonly Dictionary<(int Row, int Col), CellData> Cells = new();
    public readonly List<MergedRange> Merged = new();
    public readonly List<ColorScaleRule> ColorScales = new();
    public readonly Dictionary<(int Row, int Col), uint> ConditionalFills = new();
    public readonly List<SheetImage> Images = new();
    public readonly Dictionary<int, RowInfo> Rows = new();
    public readonly Dictionary<int, ColInfo> Cols = new();
    public double DefaultRowHeightPt = 15;
    public double DefaultColWidthChars = 8.43;
    public int MinRow = int.MaxValue;
    public int MinCol = int.MaxValue;
    public int MaxRow = 0;
    public int MaxCol = 0;

    public bool IsEmpty => MinRow == int.MaxValue || MaxRow < MinRow || MaxCol < MinCol;
}

internal enum ImageAnchorKind
{
    TwoCell,
    OneCell,
    Absolute,
}

internal sealed class SheetImage
{
    public byte[] Data = Array.Empty<byte>();
    public ImageAnchorKind Anchor = ImageAnchorKind.TwoCell;

    public int FromRow = 1;
    public int FromCol = 1;
    public long FromColOffset;
    public long FromRowOffset;

    public int ToRow = 1;
    public int ToCol = 1;
    public long ToColOffset;
    public long ToRowOffset;

    public long ExtCx;
    public long ExtCy;
    public long PosX;
    public long PosY;

    public double RotationDegrees;
    public bool FlipHorizontal;
    public bool FlipVertical;

    // 图片裁剪比例（0 - 1）
    public double CropLeft;
    public double CropTop;
    public double CropRight;
    public double CropBottom;
}


internal sealed class ColorScaleStop
{
    public string Type = "min";
    public double? Value;
    public uint Color;
}

internal sealed class ColorScaleRule
{
    public readonly List<MergedRange> Ranges = new();
    public readonly List<ColorScaleStop> Stops = new();
    public int Priority;
}


internal sealed class CliException : Exception
{
    public CliException(string message) : base(message)
    {
    }
}
