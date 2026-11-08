using System.Globalization;
using System.Text;
using Excel2Img;

internal static class Program
{
    private static int Main(string[] args)
    {
        try
        {
            Console.OutputEncoding = Encoding.UTF8;
        }
        catch (IOException)
        {
            // 某些重定向环境下无法设置编码，忽略即可
        }

        if (args.Length == 0)
        {
            PrintUsage();
            return 1;
        }

        try
        {
            var options = CliOptions.Parse(args);
            if (options.ShowHelp)
            {
                PrintUsage();
                return 0;
            }

            var (workbook, sheet) = XlsxReader.Load(options.InputPath, options.Sheet);
            using var renderer = new SheetRenderer(
                workbook,
                sheet,
                options.Scale,
                options.ShowGridLines ?? sheet.ShowGridLines);

            var (width, height) = renderer.RenderTo(options.OutputPath);
            Console.WriteLine($"已生成图片: {Path.GetFullPath(options.OutputPath)} ({width}x{height})");
            return 0;
        }
        catch (CliException ex)
        {
            Console.Error.WriteLine($"错误: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"发生未预期错误: {ex.Message}");
            return 1;
        }
    }

    private static void PrintUsage()
    {
        Console.WriteLine("""
            excel2img - 将 Excel 文件 (.xlsx) 转换为图片 (.png / .jpg)

            用法:
              excel2img <输入文件.xlsx> -o <输出文件.png|.jpg> [选项]

            选项:
              -o, --output <路径>       输出图片路径（必填），扩展名决定图片格式
              -s, --sheet <名称|序号>    指定工作表，默认使用第一个可见工作表
                  --scale <倍率>        渲染缩放倍率，默认 1；例如 2 表示输出 2 倍分辨率
                  --no-gridlines       不绘制网格线
              -h, --help                显示本帮助

            示例:
              excel2img "报表.xlsx" -o "报表.png"
              excel2img 报表.xlsx -o out.jpg --scale 2 --sheet 销售数据
            """);
    }
}

internal sealed class CliOptions
{
    public string InputPath { get; private set; } = "";
    public string OutputPath { get; private set; } = "";
    public string? Sheet { get; private set; }
    public double Scale { get; private set; } = 1.0;
    public bool? ShowGridLines { get; private set; }
    public bool ShowHelp { get; private set; }

    public static CliOptions Parse(string[] args)
    {
        var options = new CliOptions();

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            switch (arg)
            {
                case "-h":
                case "--help":
                    options.ShowHelp = true;
                    return options;
                case "-o":
                case "--output":
                    options.OutputPath = RequireValue(args, ref i, arg);
                    break;
                case "-s":
                case "--sheet":
                    options.Sheet = RequireValue(args, ref i, arg);
                    break;
                case "--scale":
                    var scaleText = RequireValue(args, ref i, arg);
                    if (!double.TryParse(scaleText, NumberStyles.Float, CultureInfo.InvariantCulture, out var scale))
                    {
                        throw new CliException($"无效的缩放倍率: {scaleText}");
                    }

                    if (scale < 0.1 || scale > 8)
                    {
                        throw new CliException("缩放倍率必须介于 0.1 和 8 之间。");
                    }

                    options.Scale = scale;
                    break;
                case "--no-gridlines":
                    options.ShowGridLines = false;
                    break;
                case "--gridlines":
                    options.ShowGridLines = true;
                    break;
                default:
                    if (arg.StartsWith("-", StringComparison.Ordinal) && arg.Length > 1)
                    {
                        throw new CliException($"未知选项: {arg}（使用 --help 查看帮助）");
                    }

                    if (options.InputPath.Length > 0)
                    {
                        throw new CliException("只能指定一个输入文件。");
                    }

                    options.InputPath = arg;
                    break;
            }
        }

        if (options.InputPath.Length == 0)
        {
            throw new CliException("缺少输入文件。用法: excel2img <输入文件.xlsx> -o <输出文件.png|.jpg>");
        }

        if (options.OutputPath.Length == 0)
        {
            throw new CliException("缺少输出文件，请使用 -o 指定输出图片路径。");
        }

        var extension = Path.GetExtension(options.OutputPath).ToLowerInvariant();
        if (extension is not (".png" or ".jpg" or ".jpeg"))
        {
            throw new CliException("输出文件扩展名必须是 .png、.jpg 或 .jpeg。");
        }

        return options;
    }

    private static string RequireValue(string[] args, ref int index, string option)
    {
        if (index + 1 >= args.Length)
        {
            throw new CliException($"选项 {option} 缺少参数值。");
        }

        index++;
        return args[index];
    }
}
