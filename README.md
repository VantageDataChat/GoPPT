# GoPPT

[English](#english) | [中文](#中文)

---

<a id="english"></a>

## English

A pure Go library for creating, reading, and writing PowerPoint (.pptx) files. Zero external dependencies. Inspired by [PHPOffice/PHPPresentation](https://github.com/PHPOffice/PHPPresentation).

### Rendering

GoPPT includes a built-in slide renderer that produces PNG images closely matching Microsoft PowerPoint's native rendering. The renderer features:

- Dual font-face measurement system (HintingNone for layout, HintingFull for rendering) to match PowerPoint's DirectWrite text metrics
- CJK-aware line wrapping with kinsoku (禁則処理) punctuation handling
- Accurate text box layout with auto-fit, auto-shrink, and overflow control
- Shape rendering: fills, borders, shadows, custom geometry paths, arrowheads
- Chart rendering: bar, line, area, pie, doughnut, scatter, radar
- Image compositing with rotation, flip, and group transforms

### Features

- Create and save `.pptx` files (OOXML / PowerPoint 2007+)
- Read existing `.pptx` files with full round-trip support
- Rich text with fonts, colors, bold, italic, underline, strikethrough
- Images (PNG, JPEG, GIF, BMP, SVG) from bytes or file path
- Tables with cell formatting and fills
- Auto shapes (rectangle, ellipse, triangle, arrows, stars, etc.)
- Line shapes with style and color
- Charts: Bar, Bar3D, Line, Area, Pie, Pie3D, Doughnut, Scatter, Radar
- Group shapes and Placeholder shapes
- Bullets (character and numeric)
- Comments with authors
- Speaker notes
- Slide backgrounds (solid and gradient)
- Animations (basic grouping)
- Document properties and custom properties
- Multiple slide layouts (4:3, 16:9, 16:10, A4, Letter, custom)
- 96% test coverage

### Installation

```bash
go get github.com/VantageDataChat/GoPPT
```

### Quick Start

```go
package main

import (
    "log"
    ppt "github.com/VantageDataChat/GoPPT"
)

func main() {
    // Create a new presentation
    p := ppt.New()

    // Set document properties
    p.GetDocumentProperties().Title = "My Presentation"
    p.GetDocumentProperties().Creator = "GoPPT"

    // First slide (created automatically)
    slide := p.GetActiveSlide()

    // Add a title
    title := slide.CreateRichTextShape()
    title.SetOffsetX(500000).SetOffsetY(300000)
    title.SetWidth(8000000).SetHeight(1000000)
    tr := title.CreateTextRun("Hello, GoPPT!")
    tr.GetFont().SetSize(28).SetBold(true).SetColor(ppt.ColorBlue)

    // Add a subtitle
    subtitle := slide.CreateRichTextShape()
    subtitle.SetOffsetX(500000).SetOffsetY(1500000)
    subtitle.SetWidth(8000000).SetHeight(600000)
    subtitle.CreateTextRun("Pure Go PowerPoint library")

    // Second slide with a chart
    slide2 := p.CreateSlide()
    chart := slide2.CreateChartShape()
    chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
    chart.BaseShape.SetWidth(7000000).SetHeight(4500000)
    chart.GetTitle().SetText("Sales Report")

    bar := ppt.NewBarChart()
    bar.AddSeries(ppt.NewChartSeriesOrdered("Revenue",
        []string{"Q1", "Q2", "Q3", "Q4"},
        []float64{120, 180, 150, 210},
    ))
    chart.GetPlotArea().SetType(bar)

    // Save
    w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
    if err := w.(*ppt.PPTXWriter).Save("presentation.pptx"); err != nil {
        log.Fatal(err)
    }
}
```

### Reading a Presentation

```go
reader := &ppt.PPTXReader{}
pres, err := reader.Read("presentation.pptx")
if err != nil {
    log.Fatal(err)
}

for i, slide := range pres.GetAllSlides() {
    fmt.Printf("Slide %d: %d shapes\n", i+1, len(slide.GetShapes()))
}
```

### Writing to io.Writer

```go
var buf bytes.Buffer
w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
w.WriteTo(&buf)
// buf.Bytes() contains the .pptx data
```

### More Examples

See [API Documentation](API.md) for the full reference, or check `example_test.go` for a working example.

---

<a id="中文"></a>

## 中文

纯 Go 语言实现的 PowerPoint (.pptx) 文件创建、读取和写入库。零外部依赖。灵感来自 [PHPOffice/PHPPresentation](https://github.com/PHPOffice/PHPPresentation)。

### 渲染能力

GoPPT 内置幻灯片渲染器，可将幻灯片导出为 PNG 图片，渲染效果接近 Microsoft PowerPoint 原生渲染。渲染器特性：

- 双字体度量系统（HintingNone 用于排版，HintingFull 用于渲染），匹配 PowerPoint DirectWrite 的文本度量
- CJK 感知的自动换行，支持禁則処理（行首行尾标点规则）
- 精确的文本框排版：自动适应、自动缩放、溢出控制
- 形状渲染：填充、边框、阴影、自定义几何路径、箭头
- 图表渲染：柱状图、折线图、面积图、饼图、环形图、散点图、雷达图
- 图片合成：旋转、翻转、组合变换

### 功能特性

- 创建和保存 `.pptx` 文件（OOXML / PowerPoint 2007+）
- 读取现有 `.pptx` 文件，支持完整的读写往返
- 富文本：字体、颜色、粗体、斜体、下划线、删除线
- 图片（PNG、JPEG、GIF、BMP、SVG），支持字节数据或文件路径
- 表格，支持单元格格式和填充
- 自动形状（矩形、椭圆、三角形、箭头、星形等）
- 线条形状，支持样式和颜色
- 图表：柱状图、3D柱状图、折线图、面积图、饼图、3D饼图、环形图、散点图、雷达图
- 组合形状和占位符形状
- 项目符号（字符和数字编号）
- 批注（含作者信息）
- 演讲者备注
- 幻灯片背景（纯色和渐变）
- 动画（基础分组）
- 文档属性和自定义属性
- 多种幻灯片布局（4:3、16:9、16:10、A4、Letter、自定义）
- 96% 测试覆盖率

### 安装

```bash
go get github.com/VantageDataChat/GoPPT
```

### 快速开始

```go
package main

import (
    "log"
    ppt "github.com/VantageDataChat/GoPPT"
)

func main() {
    // 创建新演示文稿
    p := ppt.New()

    // 设置文档属性
    p.GetDocumentProperties().Title = "我的演示文稿"
    p.GetDocumentProperties().Creator = "GoPPT"

    // 第一张幻灯片（自动创建）
    slide := p.GetActiveSlide()

    // 添加标题
    title := slide.CreateRichTextShape()
    title.SetOffsetX(500000).SetOffsetY(300000)
    title.SetWidth(8000000).SetHeight(1000000)
    tr := title.CreateTextRun("你好，GoPPT！")
    tr.GetFont().SetSize(28).SetBold(true).SetColor(ppt.ColorBlue)

    // 添加副标题
    subtitle := slide.CreateRichTextShape()
    subtitle.SetOffsetX(500000).SetOffsetY(1500000)
    subtitle.SetWidth(8000000).SetHeight(600000)
    subtitle.CreateTextRun("纯 Go 语言 PowerPoint 库")

    // 第二张幻灯片：图表
    slide2 := p.CreateSlide()
    chart := slide2.CreateChartShape()
    chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
    chart.BaseShape.SetWidth(7000000).SetHeight(4500000)
    chart.GetTitle().SetText("销售报告")

    bar := ppt.NewBarChart()
    bar.AddSeries(ppt.NewChartSeriesOrdered("营收",
        []string{"第一季度", "第二季度", "第三季度", "第四季度"},
        []float64{120, 180, 150, 210},
    ))
    chart.GetPlotArea().SetType(bar)

    // 保存
    w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
    if err := w.(*ppt.PPTXWriter).Save("演示文稿.pptx"); err != nil {
        log.Fatal(err)
    }
}
```

### 读取演示文稿

```go
reader := &ppt.PPTXReader{}
pres, err := reader.Read("演示文稿.pptx")
if err != nil {
    log.Fatal(err)
}

for i, slide := range pres.GetAllSlides() {
    fmt.Printf("幻灯片 %d：%d 个形状\n", i+1, len(slide.GetShapes()))
}
```

### 写入到 io.Writer

```go
var buf bytes.Buffer
w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
w.WriteTo(&buf)
// buf.Bytes() 包含 .pptx 数据
```

### 更多示例

请参阅 [API 文档](API.md) 获取完整参考，或查看 `example_test.go` 获取可运行的示例。

---

## License

MIT
