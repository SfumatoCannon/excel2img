using System.Globalization;
using SkiaSharp;

namespace Excel2Img;

/// <summary>
/// 字体解析：优先按完整字体名匹配，匹配不到时按名称中的样式词（Bold/Condensed/Light 等）
/// 拆出家族名再匹配，从而支持 "Tw Cen MT Condensed Extra Bold" 这类由家族名加样式词组成的字体。
/// </summary>
internal static class FontResolver
{
    private static readonly Dictionary<string, SKTypeface> Cache = new(StringComparer.Ordinal);

    private static readonly string[] CjkFonts = { "Microsoft YaHei", "SimSun", "SimHei" };

    public static SKTypeface Resolve(string? fontName, bool bold, bool italic)
    {
        var name = (fontName ?? "").Trim();
        var key = string.Create(CultureInfo.InvariantCulture, $"n|{name.ToLowerInvariant()}|{bold}|{italic}");
        if (Cache.TryGetValue(key, out var cached))
        {
            return cached;
        }

        var traits = ParseTraits(name, bold, italic);
        var result = FindBestMatch(name, traits) ?? SKTypeface.Default;
        Cache[key] = result;
        return result;
    }

    public static SKTypeface ResolveCjk(bool bold, bool italic)
    {
        var key = string.Create(CultureInfo.InvariantCulture, $"cjk|{bold}|{italic}");
        if (Cache.TryGetValue(key, out var cached))
        {
            return cached;
        }

        var weight = bold ? 700 : 400;
        var slant = italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright;
        var style = new SKFontStyle(weight, (int)SKFontStyleWidth.Normal, slant);

        SKTypeface? result = null;
        foreach (var name in CjkFonts)
        {
            result = SKFontManager.Default.MatchFamily(name, style);
            if (result != null)
            {
                break;
            }
        }

        result ??= SKTypeface.Default;
        Cache[key] = result;
        return result;
    }

    public static bool SupportsCjk(SKTypeface typeface, string text)
    {
        var checkedCount = 0;
        foreach (var rune in text.EnumerateRunes())
        {
            if (rune.Value < 0x2E80)
            {
                continue;
            }

            if (typeface.GetGlyph(rune.Value) == 0)
            {
                return false;
            }

            if (++checkedCount >= 5)
            {
                break;
            }
        }

        return true;
    }

    private static SKTypeface? FindBestMatch(string name, FontTraits traits)
    {
        if (name.Length == 0)
        {
            return SKTypeface.Default;
        }

        var words = name.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        SKTypeface? fallback = null;

        // 从完整名称开始，逐步去掉末尾的样式词，直到匹配到字体
        for (var length = words.Length; length >= 1; length--)
        {
            var candidate = string.Join(' ', words.Take(length));
            var typeface = SKFontManager.Default.MatchFamily(candidate, CreateStyle(traits));
            if (typeface == null)
            {
                continue;
            }

            if (IsExactMatch(typeface, traits))
            {
                return typeface;
            }

            fallback ??= typeface;
        }

        return fallback;
    }

    private static SKFontStyle CreateStyle(FontTraits traits) =>
        new(traits.Weight, traits.Width, traits.Slant);

    private static bool IsExactMatch(SKTypeface typeface, FontTraits traits) =>
        typeface.FontWeight == traits.Weight
        && typeface.FontWidth == traits.Width
        && typeface.FontSlant == traits.Slant;

    private static FontTraits ParseTraits(string name, bool bold, bool italic)
    {
        var lower = name.ToLowerInvariant();
        var traits = new FontTraits();

        if (ContainsAny(lower, "extra bold", "extrabold", "ultra bold", "ultrabold"))
        {
            traits.Weight = 800;
        }
        else if (ContainsAny(lower, "semi bold", "semibold", "demi bold", "demibold"))
        {
            traits.Weight = 600;
        }
        else if (ContainsAny(lower, "black", "heavy"))
        {
            traits.Weight = 900;
        }
        else if (lower.Contains("bold", StringComparison.Ordinal))
        {
            traits.Weight = 700;
        }
        else if (lower.Contains("medium", StringComparison.Ordinal))
        {
            traits.Weight = 500;
        }
        else if (ContainsAny(lower, "extra light", "extralight", "ultra light", "ultralight"))
        {
            traits.Weight = 200;
        }
        else if (lower.Contains("light", StringComparison.Ordinal))
        {
            traits.Weight = 300;
        }
        else if (lower.Contains("thin", StringComparison.Ordinal))
        {
            traits.Weight = 100;
        }

        if (bold && traits.Weight < 700)
        {
            traits.Weight = 700;
        }

        if (ContainsAny(lower, "extra condensed", "ultra condensed", "extracondensed"))
        {
            traits.Width = 2;
        }
        else if (ContainsAny(lower, "semi condensed", "semicondensed"))
        {
            traits.Width = 4;
        }
        else if (ContainsAny(lower, "condensed", "narrow"))
        {
            traits.Width = 3;
        }
        else if (ContainsAny(lower, "expanded", "wide"))
        {
            traits.Width = 7;
        }

        if (lower.Contains("oblique", StringComparison.Ordinal))
        {
            traits.Slant = SKFontStyleSlant.Oblique;
        }
        else if (lower.Contains("italic", StringComparison.Ordinal) || italic)
        {
            traits.Slant = SKFontStyleSlant.Italic;
        }

        return traits;
    }

    private static bool ContainsAny(string text, params string[] values) =>
        values.Any(value => text.Contains(value, StringComparison.Ordinal));

    private sealed class FontTraits
    {
        public int Weight = 400;
        public int Width = (int)SKFontStyleWidth.Normal;
        public SKFontStyleSlant Slant = SKFontStyleSlant.Upright;
    }
}
