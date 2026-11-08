# excel2img

[English](README.en-US.md) | 简体中文

将 Excel 文件（`.xlsx`）渲染为图片（`.png` / `.jpg`）的命令行工具。

程序直接解析 xlsx（ZIP + XML），使用 SkiaSharp 绘制表格，无需安装 Microsoft Office 或 LibreOffice。

## 主要特性

- 输出 PNG / JPEG，扩展名自动决定图片格式
- 单元格字体、字号、粗体、斜体、下划线、删除线、字体颜色
- 填充色、四边边框（含合并单元格的外边框组合）、水平/垂直对齐、缩进、自动换行
- 合并单元格、隐藏行/列、列宽、行高、网格线
- 浮动图片（twoCellAnchor / oneCellAnchor / absoluteAnchor），支持位置、尺寸、旋转、翻转与裁剪
- 列宽按工作簿默认字体的“最大数字宽度”换算，与 Excel 显示宽度一致
- 常见数字格式：常规、千分位、小数位、百分比、货币、科学计数法
- 日期与时间格式（内置与常用自定义格式，支持 1900/1904 日期系统）
- 主题色、索引色及 tint 调整
- 条件格式中的**色阶**（2 色 / 3 色，min/max/num/percent/percentile，规则优先级）
- 共享字符串、内联字符串、布尔值、错误值、公式缓存值
- 字体回退：指定字体缺失或缺少中文字形时自动回退到雅黑/宋体/黑体

## 环境要求

- 运行：Windows x64。发布的 `bin/excel2img.exe` 为单文件自包含程序，无需安装 .NET 运行时
- 构建：.NET 10 SDK（Windows）

## 构建

```powershell
dotnet publish -c Release -o bin
```

输出：`bin/excel2img.exe`（单文件，约 40 MB）。

也可以直接双击或运行项目根目录下的 `build.bat`，它会自动执行上述命令并清理中间输出目录。

## 用法

```text
excel2img <输入文件.xlsx> -o <输出文件.png|.jpg> [选项]
```

### 选项

| 选项 | 说明 |
| --- | --- |
| `-o, --output <路径>` | 输出图片路径（必填），扩展名决定格式：`.png`、`.jpg`、`.jpeg` |
| `-s, --sheet <名称\|序号>` | 指定工作表；序号从 1 开始。默认使用第一个可见工作表 |
| `--scale <倍率>` | 渲染缩放倍率，默认 `1`，范围 `0.1` – `8`。例如 `2` 输出 2 倍分辨率 |
| `--no-gridlines` | 不绘制网格线 |
| `--gridlines` | 强制绘制网格线（默认跟随工作表设置） |
| `-h, --help` | 显示帮助 |

### 示例

```powershell
# 基本用法
bin\excel2img.exe "报表.xlsx" -o "报表.png"

# 输出 JPG，2 倍分辨率
bin\excel2img.exe "报表.xlsx" -o "报表.jpg" --scale 2

# 指定工作表（名称或序号）
bin\excel2img.exe "报表.xlsx" -o "明细.png" --sheet "明细"
bin\excel2img.exe "报表.xlsx" -o "明细.png" --sheet 2

# 不绘制网格线
bin\excel2img.exe "报表.xlsx" -o "报表.png" --no-gridlines
```

### 退出码

| 退出码 | 含义 |
| --- | --- |
| `0` | 成功 |
| `1` | 参数错误、文件不存在、格式不支持或渲染失败 |

## 界面语言

命令行帮助、提示与错误信息支持英文和中文，**默认英文**。语言由当前工作目录或可执行文件附近的 `.env` 文件控制：

```ini
# en-US（默认）或 zh-CN
EXCEL2IMG_LANG=en-US
```

- 可识别的键：`EXCEL2IMG_LANG`（别名 `EXCEL2IMG_LANGUAGE`）
- `.env` 查找顺序：可执行文件所在目录 → 当前工作目录 → 可执行文件的上层目录（最多 5 级）
- 修改 `.env` 后重新运行命令即可生效；找不到 `.env` 时使用英文
- `.env` 已被 `.gitignore` 忽略，仅作为本地配置

## 渲染说明

- 图片按 **96 DPI** 渲染，即 `1 磅 = 4/3 像素`，与 Excel 在 100% 缩放下的显示尺寸一致。
- 若希望图片像素值与 Excel 行高/列宽中的“磅”数值 1:1（例如行高 608 磅对应 608 像素），可使用 `--scale 0.75`。
- 合并单元格的外边框会从合并区域四角单元格组合（Excel 会把底边框存放在最后一行、右边框存放在最右列）。
- 自动换行内容略高于单元格时，会轻微压缩行距以避免最后一行被裁掉（最多压缩 15%）。

## 已知限制

- 仅支持 `.xlsx`，不支持旧版 `.xls`
- 不渲染图表、形状、批注、数据透视表等对象；图片支持 PNG/JPEG/GIF/BMP 等 Skia 可解码格式，
  不支持 EMF/WMF/SVG 以及“放置在单元格中”的内嵌图片（DISPIMG）
- 条件格式仅支持**色阶**，不支持数据条、图标集及公式/规则型条件格式
- 公式使用文件中保存的缓存值，不重新计算公式
- 不处理打印区域、分页符、页眉页脚等打印布局设置
- 单元格内容在“超过单元格尺寸且压缩行距后仍放不下”时仍会像 Excel 一样被裁切

## 项目结构

| 文件 | 说明 |
| --- | --- |
| `Program.cs` | 命令行参数解析与入口 |
| `XlsxReader.cs` | xlsx 解析：工作表、样式、共享字符串、合并单元格、条件格式等 |
| `NumberFormatter.cs` | 数字、日期、百分比、货币等格式化为显示文本 |
| `SheetRenderer.cs` | 布局计算与 SkiaSharp 绘制、图片编码输出 |
| `FontResolver.cs` | 字体解析、样式词拆分与中文字形回退 |
| `Localization.cs` | 界面语言加载与中英文文案 |
| `.env` | 本地语言配置（默认 `en-US`） |
| `Model.cs` | 工作簿、工作表、单元格样式等数据模型 |
| `excel2img.csproj` | 项目与发布配置（`win-x64`、单文件、自包含） |
| `build.bat` | 一键构建脚本 |

## 许可证

本项目未声明开源许可证。
