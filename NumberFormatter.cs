using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel2Img;

internal static class NumberFormatter
{
    public static string FormatValue(CellData cell, CellStyle style, Workbook workbook)
    {
        switch (cell.Kind)
        {
            case CellValueKind.Empty:
                return "";
            case CellValueKind.SharedString:
                return cell.SharedStringIndex >= 0 && cell.SharedStringIndex < workbook.SharedStrings.Count
                    ? workbook.SharedStrings[cell.SharedStringIndex]
                    : "";
            case CellValueKind.InlineString:
            case CellValueKind.String:
                return cell.RawValue;
            case CellValueKind.Bool:
                return cell.RawValue == "1" ? "TRUE" : "FALSE";
            case CellValueKind.Error:
                return cell.RawValue;
        }

        if (!double.TryParse(cell.RawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
        {
            return cell.RawValue;
        }

        var format = style.FormatCode ?? GetBuiltinFormat(style.NumFmtId);
        if (string.IsNullOrWhiteSpace(format))
        {
            return FormatGeneral(value, cell.RawValue);
        }

        try
        {
            return FormatNumber(value, format, cell.RawValue, workbook.Date1904);
        }
        catch (Exception)
        {
            return FormatGeneral(value, cell.RawValue);
        }
    }

    private static string? GetBuiltinFormat(int id) => id switch
    {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "$#,##0_);($#,##0)",
        6 => "$#,##0_);[Red]($#,##0)",
        7 => "$#,##0.00_);($#,##0.00)",
        8 => "$#,##0.00_);[Red]($#,##0.00)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "yyyy/m/d",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "yyyy/m/d h:mm",
        37 => "#,##0_);(#,##0)",
        38 => "#,##0_);[Red](#,##0)",
        39 => "#,##0.00_);(#,##0.00)",
        40 => "#,##0.00_);[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mm:ss.0",
        48 => "##0.0E+0",
        49 => "@",
        >= 27 and <= 36 => "yyyy/m/d",
        >= 50 and <= 58 => "yyyy/m/d",
        _ => null,
    };

    private static string FormatNumber(double value, string format, string rawValue, bool date1904)
    {
        if (string.Equals(format, "General", StringComparison.OrdinalIgnoreCase) || format.Length == 0)
        {
            return FormatGeneral(value, rawValue);
        }

        if (format.Trim() == "@")
        {
            return rawValue;
        }

        var sections = SplitSections(format);
        var sectionIndex = 0;
        var usingNegativeSection = false;
        if (value < 0 && sections.Count >= 2 && !string.IsNullOrWhiteSpace(sections[1]))
        {
            sectionIndex = 1;
            usingNegativeSection = true;
        }
        else if (value == 0 && sections.Count >= 3 && !string.IsNullOrWhiteSpace(sections[2]))
        {
            sectionIndex = 2;
            usingNegativeSection = true;
        }

        var section = sections[sectionIndex];
        if (IsDateFormat(section))
        {
            return FormatDate(value, section, date1904);
        }

        return FormatNumeric(value, section, value < 0 && !usingNegativeSection);
    }

    private static string FormatGeneral(double value, string rawValue)
    {
        if (string.IsNullOrWhiteSpace(rawValue))
        {
            return "";
        }

        try
        {
            return value.ToString("G11", CultureInfo.InvariantCulture);
        }
        catch (FormatException)
        {
            return rawValue;
        }
    }

    private static string FormatNumeric(double value, string section, bool addMinus)
    {
        var magnitude = Math.Abs(value);

        var hasPercent = section.Contains('%');
        if (hasPercent)
        {
            magnitude *= 100;
        }

        var placeholderStart = FindPlaceholderStart(section);
        var placeholderEnd = FindPlaceholderEnd(section);
        var placeholder = placeholderStart >= 0 && placeholderEnd >= placeholderStart
            ? section[placeholderStart..(placeholderEnd + 1)]
            : "0";

        var scientific = placeholder.Contains('E') || placeholder.Contains('e');
        var groupSeparator = Regex.IsMatch(placeholder, @"[0#?],[0#?]");

        var dotIndex = placeholder.IndexOf('.');
        var decimalPart = dotIndex >= 0 ? placeholder[(dotIndex + 1)..] : "";
        var decimals = decimalPart.Count(c => c is '0' or '#' or '?');
        var mandatoryDecimals = decimalPart.TakeWhile(c => c == '0').Count();

        string number;
        if (scientific)
        {
            number = FormatScientific(magnitude, decimals);
        }
        else
        {
            var numberFormat = groupSeparator
                ? (decimals > 0 ? "N" + decimals : "N0")
                : (decimals > 0 ? "F" + decimals : "F0");
            number = magnitude.ToString(numberFormat, CultureInfo.InvariantCulture);

            if (decimals > mandatoryDecimals && number.Contains('.'))
            {
                var trimmed = number.TrimEnd('0');
                var dot = trimmed.IndexOf('.');
                if (dot >= 0)
                {
                    var keep = Math.Min(trimmed.Length, dot + 1 + mandatoryDecimals);
                    trimmed = trimmed[..keep];
                    trimmed = trimmed.TrimEnd('.');
                }

                number = trimmed;
            }
        }

        CollectLiterals(section, out var prefix, out var suffix, hasPercent);

        var result = prefix + number + suffix;
        if (hasPercent && !suffix.Contains('%'))
        {
            result += "%";
        }

        if (addMinus)
        {
            result = "-" + result;
        }

        return result;
    }

    private static string FormatScientific(double value, int decimals)
    {
        if (value == 0)
        {
            return "0" + (decimals > 0 ? "." + new string('0', decimals) : "") + "E+00";
        }

        var exponent = (int)Math.Floor(Math.Log10(value));
        var mantissa = value / Math.Pow(10, exponent);
        if (Math.Round(mantissa, decimals) >= 10)
        {
            mantissa /= 10;
            exponent++;
        }

        var pattern = "0" + (decimals > 0 ? "." + new string('0', decimals) : "");
        return mantissa.ToString(pattern, CultureInfo.InvariantCulture)
               + "E"
               + (exponent >= 0 ? "+" : "-")
               + Math.Abs(exponent).ToString("00", CultureInfo.InvariantCulture);
    }

    private static void CollectLiterals(string section, out string prefix, out string suffix, bool hasPercent)
    {
        var prefixBuilder = new StringBuilder();
        var suffixBuilder = new StringBuilder();
        var placeholderStart = FindPlaceholderStart(section);
        var index = 0;

        while (index < section.Length)
        {
            var ch = section[index];

            if (ch == '"')
            {
                var end = section.IndexOf('"', index + 1);
                if (end < 0)
                {
                    end = section.Length;
                }

                var literal = section[(index + 1)..end];
                if (index < placeholderStart) prefixBuilder.Append(literal);
                else suffixBuilder.Append(literal);
                index = end + 1;
                continue;
            }

            if (ch == '\\')
            {
                if (index + 1 < section.Length)
                {
                    var literal = section[index + 1].ToString();
                    if (index < placeholderStart) prefixBuilder.Append(literal);
                    else suffixBuilder.Append(literal);
                }

                index += 2;
                continue;
            }

            if (ch == '_' || ch == '*')
            {
                index += 2;
                continue;
            }

            if (ch == '[')
            {
                var end = section.IndexOf(']', index + 1);
                index = end < 0 ? section.Length : end + 1;
                continue;
            }

            if (ch is '0' or '#' or '?' or '.' or ',' or 'E' or 'e' or '+' or '-'
                || (ch == '%' && hasPercent))
            {
                index++;
                continue;
            }

            if (ch == '%')
            {
                index++;
                continue;
            }

            if (index < placeholderStart || placeholderStart < 0) prefixBuilder.Append(ch);
            else suffixBuilder.Append(ch);
            index++;
        }

        prefix = prefixBuilder.ToString();
        suffix = suffixBuilder.ToString();
    }

    private static int FindPlaceholderStart(string section)
    {
        for (var i = 0; i < section.Length; i++)
        {
            if (section[i] is '0' or '#' or '?')
            {
                return i;
            }
        }

        return -1;
    }

    private static int FindPlaceholderEnd(string section)
    {
        for (var i = section.Length - 1; i >= 0; i--)
        {
            if (section[i] is '0' or '#' or '?')
            {
                return i;
            }
        }

        return -1;
    }

    private static List<string> SplitSections(string format)
    {
        var sections = new List<string>();
        var builder = new StringBuilder();
        var inQuotes = false;

        foreach (var ch in format)
        {
            if (ch == '"')
            {
                inQuotes = !inQuotes;
                builder.Append(ch);
                continue;
            }

            if (ch == ';' && !inQuotes)
            {
                sections.Add(builder.ToString());
                builder.Clear();
                continue;
            }

            builder.Append(ch);
        }

        sections.Add(builder.ToString());
        return sections;
    }

    private static bool IsDateFormat(string format)
    {
        var cleaned = StripLiterals(format);
        return cleaned.Any(c => "ymdhs".IndexOf(char.ToLowerInvariant(c)) >= 0);
    }

    private static string StripLiterals(string format)
    {
        var builder = new StringBuilder();
        var index = 0;

        while (index < format.Length)
        {
            var ch = format[index];
            if (ch == '"')
            {
                var end = format.IndexOf('"', index + 1);
                index = end < 0 ? format.Length : end + 1;
                continue;
            }

            if (ch == '\\')
            {
                index += 2;
                continue;
            }

            if (ch == '[')
            {
                var end = format.IndexOf(']', index + 1);
                index = end < 0 ? format.Length : end + 1;
                continue;
            }

            if (ch == '_' || ch == '*')
            {
                index += 2;
                continue;
            }

            builder.Append(ch);
            index++;
        }

        return builder.ToString();
    }

    private static string FormatDate(double serial, string format, bool date1904)
    {
        var durationMatch = Regex.Match(format, @"\[(h+|m+|s+)\]", RegexOptions.IgnoreCase);
        if (durationMatch.Success)
        {
            return FormatDuration(serial, durationMatch.Groups[1].Value);
        }

        DateTime dateTime;
        try
        {
            dateTime = SerialToDateTime(serial, date1904);
        }
        catch (ArgumentException)
        {
            return serial.ToString(CultureInfo.InvariantCulture);
        }

        var builder = new StringBuilder();
        var hasAmPm = Regex.IsMatch(format, "AM/PM|A/P", RegexOptions.IgnoreCase);
        var index = 0;

        while (index < format.Length)
        {
            var ch = format[index];

            if (ch == '"')
            {
                var end = format.IndexOf('"', index + 1);
                if (end < 0)
                {
                    end = format.Length;
                }

                builder.Append(format[(index + 1)..end]);
                index = end + 1;
                continue;
            }

            if (ch == '\\')
            {
                if (index + 1 < format.Length)
                {
                    builder.Append(format[index + 1]);
                }

                index += 2;
                continue;
            }

            if (ch == '_' || ch == '*')
            {
                index += 2;
                continue;
            }

            if (ch == '[')
            {
                var end = format.IndexOf(']', index + 1);
                index = end < 0 ? format.Length : end + 1;
                continue;
            }

            var lower = char.ToLowerInvariant(ch);
            if (lower is 'y' or 'm' or 'd' or 'h' or 's')
            {
                var run = 1;
                while (index + run < format.Length && char.ToLowerInvariant(format[index + run]) == lower)
                {
                    run++;
                }

                switch (lower)
                {
                    case 'y':
                        builder.Append(run >= 4
                            ? dateTime.Year.ToString("0000", CultureInfo.InvariantCulture)
                            : (dateTime.Year % 100).ToString("00", CultureInfo.InvariantCulture));
                        break;
                    case 'm':
                        if (IsMinuteToken(format, index, run))
                        {
                            builder.Append(run >= 2
                                ? dateTime.Minute.ToString("00", CultureInfo.InvariantCulture)
                                : dateTime.Minute.ToString(CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            builder.Append(run switch
                            {
                                >= 4 => dateTime.ToString("MMMM", CultureInfo.CurrentCulture),
                                3 => dateTime.ToString("MMM", CultureInfo.CurrentCulture),
                                2 => dateTime.Month.ToString("00", CultureInfo.InvariantCulture),
                                _ => dateTime.Month.ToString(CultureInfo.InvariantCulture),
                            });
                        }

                        break;
                    case 'd':
                        builder.Append(run switch
                        {
                            >= 4 => dateTime.ToString("dddd", CultureInfo.CurrentCulture),
                            3 => dateTime.ToString("ddd", CultureInfo.CurrentCulture),
                            2 => dateTime.Day.ToString("00", CultureInfo.InvariantCulture),
                            _ => dateTime.Day.ToString(CultureInfo.InvariantCulture),
                        });
                        break;
                    case 'h':
                        if (hasAmPm)
                        {
                            var hour = dateTime.Hour % 12;
                            if (hour == 0)
                            {
                                hour = 12;
                            }

                            builder.Append(run >= 2
                                ? hour.ToString("00", CultureInfo.InvariantCulture)
                                : hour.ToString(CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            builder.Append(run >= 2
                                ? dateTime.Hour.ToString("00", CultureInfo.InvariantCulture)
                                : dateTime.Hour.ToString(CultureInfo.InvariantCulture));
                        }

                        break;
                    case 's':
                        builder.Append(run >= 2
                            ? dateTime.Second.ToString("00", CultureInfo.InvariantCulture)
                            : dateTime.Second.ToString(CultureInfo.InvariantCulture));
                        break;
                }

                index += run;
                continue;
            }

            if (lower == 'a')
            {
                var rest = format[index..];
                if (rest.StartsWith("AM/PM", StringComparison.OrdinalIgnoreCase))
                {
                    builder.Append(dateTime.Hour < 12 ? "AM" : "PM");
                    index += 5;
                    continue;
                }

                if (rest.StartsWith("A/P", StringComparison.OrdinalIgnoreCase))
                {
                    builder.Append(dateTime.Hour < 12 ? "A" : "P");
                    index += 3;
                    continue;
                }
            }

            builder.Append(ch);
            index++;
        }

        return builder.ToString();
    }

    private static bool IsMinuteToken(string format, int index, int run)
    {
        for (var i = index - 1; i >= 0; i--)
        {
            var c = format[i];
            if (c == '"')
            {
                var start = format.LastIndexOf('"', i - 1);
                if (start < 0)
                {
                    break;
                }

                i = start;
                continue;
            }

            if (c == ']')
            {
                var start = format.LastIndexOf('[', i);
                if (start < 0)
                {
                    break;
                }

                i = start;
                continue;
            }

            if (char.IsLetter(c))
            {
                return char.ToLowerInvariant(c) == 'h';
            }
        }

        for (var i = index + run; i < format.Length; i++)
        {
            var c = format[i];
            if (c == '"')
            {
                var end = format.IndexOf('"', i + 1);
                if (end < 0)
                {
                    break;
                }

                i = end;
                continue;
            }

            if (c == '[')
            {
                var end = format.IndexOf(']', i + 1);
                if (end < 0)
                {
                    break;
                }

                i = end;
                continue;
            }

            if (char.IsLetter(c))
            {
                return char.ToLowerInvariant(c) == 's';
            }
        }

        return false;
    }

    private static string FormatDuration(double serial, string unit)
    {
        var totalSeconds = Math.Abs(serial) * 86400.0;
        var seconds = (long)Math.Round(totalSeconds);
        var hours = seconds / 3600;
        var minutes = (seconds % 3600) / 60;
        var secs = seconds % 60;
        var sign = serial < 0 ? "-" : "";

        return unit.ToLowerInvariant()[0] switch
        {
            'h' => $"{sign}{hours}:{minutes:00}:{secs:00}",
            'm' => $"{sign}{hours * 60 + minutes}:{secs:00}",
            _ => $"{sign}{seconds}",
        };
    }

    private static DateTime SerialToDateTime(double serial, bool date1904)
    {
        if (date1904)
        {
            return new DateTime(1904, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).AddDays(serial);
        }

        return serial < 60
            ? DateTime.FromOADate(serial + 1)
            : DateTime.FromOADate(serial);
    }
}
