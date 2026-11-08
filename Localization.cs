using System.Globalization;

namespace Excel2Img;

/// <summary>
/// 界面语言设置：默认英文，可通过 .env 中的 EXCEL2IMG_LANG 切换为中文。
/// .env 会从当前工作目录、可执行文件所在目录及其上层目录（最多 5 级）中查找。
/// </summary>
internal static class Localization
{
    private const string LanguageKey = "EXCEL2IMG_LANG";
    private const string LanguageKeyAlias = "EXCEL2IMG_LANGUAGE";

    public static bool UseChinese { get; private set; }

    public static void Load()
    {
        foreach (var path in EnumerateEnvFiles())
        {
            var language = TryReadLanguage(path);
            if (language == null)
            {
                continue;
            }

            UseChinese = language.StartsWith("zh", StringComparison.OrdinalIgnoreCase);
            return;
        }
    }

    public static string T(string english, string chinese) => UseChinese ? chinese : english;

    public static string F(string english, string chinese, params object?[] args) =>
        string.Format(CultureInfo.InvariantCulture, UseChinese ? chinese : english, args);

    private static IEnumerable<string> EnumerateEnvFiles()
    {
        var candidates = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void AddDirectory(string? directory)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                return;
            }

            var path = Path.Combine(directory, ".env");
            if (seen.Add(path))
            {
                candidates.Add(path);
            }
        }

        // 优先使用可执行文件旁的配置（便于分发时每个目录单独设置），
        // 其次是当前工作目录，最后回退到可执行文件的上层目录（便于开发/构建目录）。
        var executableDirectory = GetExecutableDirectory();
        AddDirectory(executableDirectory);
        AddDirectory(Environment.CurrentDirectory);

        if (executableDirectory != null)
        {
            var current = new DirectoryInfo(executableDirectory).Parent;
            for (var level = 0; level < 5 && current != null; level++, current = current.Parent)
            {
                AddDirectory(current.FullName);
            }
        }

        return candidates;
    }

    private static string? GetExecutableDirectory()
    {
        try
        {
            var processPath = Environment.ProcessPath;
            if (!string.IsNullOrEmpty(processPath))
            {
                return Path.GetDirectoryName(processPath);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // 忽略，回退到 BaseDirectory
        }

        return AppContext.BaseDirectory;
    }

    private static string? TryReadLanguage(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return null;
            }

            string? language = null;
            foreach (var rawLine in File.ReadLines(path))
            {
                var line = rawLine.TrimStart('\uFEFF').Trim();
                if (line.Length == 0 || line.StartsWith('#'))
                {
                    continue;
                }

                var separator = line.IndexOf('=');
                if (separator <= 0)
                {
                    continue;
                }

                var key = line[..separator].Trim();
                if (!key.Equals(LanguageKey, StringComparison.OrdinalIgnoreCase)
                    && !key.Equals(LanguageKeyAlias, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var value = line[(separator + 1)..].Trim().Trim('"', '\'');
                var commentIndex = value.IndexOfAny(new[] { ' ', '\t', '#' });
                if (commentIndex >= 0)
                {
                    value = value[..commentIndex];
                }

                if (value.Length > 0)
                {
                    language = value;
                }
            }

            return language;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return null;
        }
    }
}
