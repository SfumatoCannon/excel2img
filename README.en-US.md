# excel2img

English | [简体中文](README.zh-CN.md)

A command-line tool that renders Excel workbooks (`.xlsx`) to images (`.png` / `.jpg`).

The xlsx file is parsed directly from its ZIP + XML parts and drawn with SkiaSharp. Microsoft Office or LibreOffice is **not** required.

## Features

- PNG / JPEG output; the file extension selects the format
- Font family, size, bold, italic, underline, strikethrough and font color
- Solid fills, cell borders (merged-cell outer borders are composed from edge cells),
  horizontal/vertical alignment, indent and text wrapping
- Merged cells, hidden rows/columns, column widths, row heights and gridlines
- Column widths follow the workbook's default font "maximum digit width", matching Excel's on-screen width
- Common number formats: General, thousands separator, fixed decimals, percentage,
  currency and scientific notation
- Date/time formats (built-in and common custom patterns, 1900/1904 date systems)
- Theme colors, indexed colors and tint adjustment
- **Color scales** from conditional formatting (2-color / 3-color, min/max/num/percent/percentile,
  rule priority)
- Shared strings, inline strings, booleans, error values and cached formula results
- Font fallback: if the requested font is missing or lacks CJK glyphs, Microsoft YaHei / SimSun / SimHei is used

## Requirements

- Runtime: Windows x64. The published `bin/excel2img.exe` is a self-contained single-file executable
  and does not require the .NET runtime to be installed.
- Build: .NET 10 SDK (Windows)

## Build

```powershell
dotnet publish -c Release -o bin
```

Output: `bin/excel2img.exe` (single file, about 40 MB).

## Usage

```text
excel2img <input.xlsx> -o <output.png|.jpg> [options]
```

### Options

| Option | Description |
| --- | --- |
| `-o, --output <path>` | Output image path (required). Extension selects the format: `.png`, `.jpg`, `.jpeg` |
| `-s, --sheet <name\|index>` | Worksheet to render; index is 1-based. Defaults to the first visible sheet |
| `--scale <factor>` | Render scale, default `1`, range `0.1` – `8`. For example `2` produces a 2x resolution image |
| `--no-gridlines` | Do not draw gridlines |
| `--gridlines` | Force gridlines on (default: follow the worksheet setting) |
| `-h, --help` | Show help |

### Examples

```powershell
# Basic usage
bin\excel2img.exe "report.xlsx" -o "report.png"

# JPEG output at 2x resolution
bin\excel2img.exe "report.xlsx" -o "report.jpg" --scale 2

# Select a worksheet by name or 1-based index
bin\excel2img.exe "report.xlsx" -o "details.png" --sheet "Details"
bin\excel2img.exe "report.xlsx" -o "details.png" --sheet 2

# Hide gridlines
bin\excel2img.exe "report.xlsx" -o "report.png" --no-gridlines
```

### Exit codes

| Code | Meaning |
| --- | --- |
| `0` | Success |
| `1` | Invalid arguments, missing input file, unsupported format or rendering failure |

## Rendering notes

- Images are rendered at **96 DPI**, i.e. `1 pt = 4/3 px`, matching how Excel looks at 100% zoom.
- To make pixel dimensions equal the point values shown by Excel (e.g. a 608 pt merged cell
  becoming 608 px tall), use `--scale 0.75`.
- Outer borders of merged cells are composed from the four corner cells, because Excel stores
  the bottom border on the last row and the right border on the rightmost column.
- When wrapped text is slightly taller than its cell, line spacing is compressed by up to 15%
  so that the last line is not clipped.

## Limitations

- Only `.xlsx` is supported; legacy `.xls` is not
- Images, charts, shapes, comments and pivot tables are not rendered
- Only **color scales** of conditional formatting are supported; data bars, icon sets and
  formula/rule-based conditional formatting are not
- Formulas are not recalculated; the cached values stored in the file are used
- Print areas, page breaks, headers and footers are ignored
- Content that genuinely exceeds its cell (even after line-spacing compression) is clipped,
  as it is in Excel

## Project structure

| File | Description |
| --- | --- |
| `Program.cs` | Command-line parsing and entry point |
| `XlsxReader.cs` | xlsx parsing: sheets, styles, shared strings, merged cells, conditional formatting, etc. |
| `NumberFormatter.cs` | Formatting of numbers, dates, percentages and currency into display text |
| `SheetRenderer.cs` | Layout calculation, SkiaSharp drawing and image encoding |
| `FontResolver.cs` | Font resolution, style-word splitting and CJK glyph fallback |
| `Model.cs` | Data model for workbooks, worksheets and cell styles |
| `excel2img.csproj` | Project and publish settings (`win-x64`, single file, self-contained) |

## License

No open-source license has been declared for this project.
