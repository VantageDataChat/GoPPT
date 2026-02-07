# GoPPT API Reference / API 参考文档

[English](#english) | [中文](#中文)

---

<a id="english"></a>

## English

Package: `github.com/VantageDataChat/GoPPT`

All dimensions use EMU (English Metric Units): 1 inch = 914400 EMU, 1 cm = 360000 EMU, 1 pt = 12700 EMU.

---

### Presentation

The root object representing a PowerPoint file.

```go
// Create a new presentation (includes one blank slide)
p := ppt.New()

// Document properties
p.GetDocumentProperties().Title = "Title"
p.GetDocumentProperties().Creator = "Author"
p.GetDocumentProperties().Description = "Description"
p.GetDocumentProperties().Subject = "Subject"
p.GetDocumentProperties().Keywords = "go, pptx"
p.GetDocumentProperties().Category = "Report"
p.GetDocumentProperties().Company = "ACME"

// Custom properties
p.GetDocumentProperties().SetCustomProperty("version", "1.0", ppt.PropertyTypeString)
p.GetDocumentProperties().GetCustomPropertyValue("version") // "1.0"

// Presentation properties
p.GetPresentationProperties().SetZoom(1.5)
p.GetPresentationProperties().SetLastView(ppt.ViewSlide)
p.GetPresentationProperties().SetSlideshowType(ppt.SlideshowTypePresent)
p.GetPresentationProperties().SetCommentVisible(true)
p.GetPresentationProperties().MarkAsFinal()

// Layout
p.GetLayout().SetLayout(ppt.LayoutScreen16x9)
p.GetLayout().SetCustomLayout(9144000, 6858000) // custom EMU dimensions
```

| Layout Constant | Description |
|---|---|
| `LayoutScreen4x3` | 10" × 7.5" (default) |
| `LayoutScreen16x9` | 13.33" × 7.5" |
| `LayoutScreen16x10` | 12" × 7.5" |
| `LayoutA4` | A4 landscape |
| `LayoutLetter` | US Letter |
| `LayoutCustom` | Custom dimensions |

---

### Slides

```go
slide := p.CreateSlide()           // create and append
slide := p.GetActiveSlide()        // get current active slide
p.SetActiveSlideIndex(1)           // switch active slide
slide, _ := p.GetSlide(0)         // get by index
slides := p.GetAllSlides()         // get all
count := p.GetSlideCount()         // count
p.RemoveSlideByIndex(0)            // remove

slide.SetName("Intro")
slide.SetNotes("Speaker notes here")
slide.SetVisible(true)
slide.SetBackground(ppt.NewFill().SetSolid(ppt.ColorWhite))
```

---

### Shapes

All shapes share a common `BaseShape` with position, size, fill, border, shadow, and hyperlink.

```go
// Common BaseShape methods (available on all shapes)
shape.BaseShape.SetOffsetX(914400)   // 1 inch from left
shape.BaseShape.SetOffsetY(914400)   // 1 inch from top
shape.BaseShape.SetWidth(5000000)
shape.BaseShape.SetHeight(3000000)
shape.BaseShape.SetName("My Shape")
shape.BaseShape.SetRotation(45)      // degrees
shape.BaseShape.SetFill(fill)
shape.BaseShape.SetBorder(border)
shape.BaseShape.SetShadow(shadow)
shape.BaseShape.SetHyperlink(ppt.NewHyperlink("https://example.com"))
```

#### RichTextShape

```go
rt := slide.CreateRichTextShape()
rt.SetOffsetX(100).SetOffsetY(100).SetWidth(8000000).SetHeight(1000000)
rt.SetWordWrap(true)
rt.SetAutoFit(ppt.AutoFitNormal)
rt.SetColumns(2)

// Paragraphs and text runs
para := rt.GetActiveParagraph()
tr := para.CreateTextRun("Hello World")
tr.GetFont().SetBold(true).SetSize(24).SetColor(ppt.ColorRed).SetName("Arial")
tr.GetFont().SetItalic(true).SetUnderline(ppt.UnderlineSingle).SetStrikethrough(true)

// Line break
para.CreateBreak()
para.CreateTextRun("Second line")

// New paragraph
para2 := rt.CreateParagraph()
para2.GetAlignment().SetHorizontal(ppt.HorizontalCenter)
para2.SetLineSpacing(200)
para2.SetSpaceBefore(100)
para2.SetSpaceAfter(50)
```

#### DrawingShape (Images)

```go
// From byte data
img := slide.CreateDrawingShape()
img.SetImageData(imageBytes, "image/png")
img.SetWidth(2000000).SetHeight(1500000).SetOffsetX(100).SetOffsetY(100)

// From file path
img2 := ppt.NewDrawingShape()
img2.SetPath("/path/to/image.jpg")
img2.SetWidth(2000000).SetHeight(1500000)
slide.AddShape(img2)
```

Supported formats: PNG, JPEG, GIF, BMP, SVG.

#### TableShape

```go
table := slide.CreateTableShape(3, 4) // 3 rows, 4 columns
table.SetWidth(8000000).SetHeight(2000000)
table.BaseShape.SetOffsetX(500000).SetOffsetY(2000000)

cell := table.GetCell(0, 0) // row 0, col 0
cell.SetText("Header")
cell.SetFill(ppt.NewFill().SetSolid(ppt.ColorBlue))
cell.SetColSpan(2)
cell.SetRowSpan(1)
```

#### AutoShape

```go
shape := slide.CreateAutoShape()
shape.SetAutoShapeType(ppt.AutoShapeRoundedRect)
shape.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(2000000).SetHeight(1000000)
shape.SetText("Inside the shape")
shape.BaseShape.SetFill(ppt.NewFill().SetSolid(ppt.ColorYellow))
```

| AutoShape Type | Constant |
|---|---|
| Rectangle | `AutoShapeRectangle` |
| Rounded Rectangle | `AutoShapeRoundedRect` |
| Ellipse | `AutoShapeEllipse` |
| Triangle | `AutoShapeTriangle` |
| Diamond | `AutoShapeDiamond` |
| Pentagon | `AutoShapePentagon` |
| Hexagon | `AutoShapeHexagon` |
| Star (4/5 point) | `AutoShapeStar4`, `AutoShapeStar5` |
| Arrows | `AutoShapeArrowRight/Left/Up/Down` |
| Heart | `AutoShapeHeart` |
| Lightning Bolt | `AutoShapeLightningBolt` |

#### LineShape

```go
line := slide.CreateLineShape()
line.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(0)
line.SetLineWidth(2).SetLineColor(ppt.ColorRed).SetLineStyle(ppt.BorderSolid)
```

#### GroupShape

```go
group := slide.CreateGroupShape()
group.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)

child := ppt.NewRichTextShape()
child.SetOffsetX(0).SetOffsetY(0).SetWidth(2000000).SetHeight(500000)
child.CreateTextRun("Inside group")
group.AddShape(child)

group.GetShapes()      // []Shape
group.GetShapeCount()  // int
group.RemoveShape(0)   // remove by index
```

#### PlaceholderShape

```go
ph := slide.CreatePlaceholderShape(ppt.PlaceholderTitle)
ph.BaseShape.SetOffsetX(500000).SetOffsetY(300000).SetWidth(8000000).SetHeight(1000000)
ph.CreateTextRun("Slide Title")
ph.SetPlaceholderIndex(0)
```

| Placeholder Type | Constant |
|---|---|
| Title | `PlaceholderTitle` |
| Body | `PlaceholderBody` |
| Center Title | `PlaceholderCtrTitle` |
| Subtitle | `PlaceholderSubTitle` |
| Date | `PlaceholderDate` |
| Footer | `PlaceholderFooter` |
| Slide Number | `PlaceholderSlideNum` |

---

### Charts

```go
chart := slide.CreateChartShape()
chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
chart.BaseShape.SetWidth(7000000).SetHeight(4500000)

// Title
chart.GetTitle().SetText("My Chart").SetVisible(true)
chart.GetTitle().Font.SetBold(true).SetSize(14)

// Legend
chart.GetLegend().Visible = true
chart.GetLegend().Position = ppt.LegendBottom // b, t, l, r, tr

// Display blank values
chart.SetDisplayBlankAs(ppt.ChartBlankAsZero) // "gap", "zero", "span"

// 3D view (for 3D charts)
chart.GetView3D().RotX = 15
chart.GetView3D().RotY = 20
```

#### Chart Types

```go
// Bar / Column
bar := ppt.NewBarChart()
bar.SetBarGrouping(ppt.BarGroupingClustered) // clustered, stacked, percentStacked
bar.SetGapWidthPercent(150)  // 0-500
bar.SetOverlapPercent(0)     // -100 to 100
bar.AddSeries(ppt.NewChartSeriesOrdered("Sales", categories, values))

// 3D Bar
bar3d := ppt.NewBar3DChart()

// Line
line := ppt.NewLineChart()
line.SetSmooth(true)

// Area
area := ppt.NewAreaChart()

// Pie
pie := ppt.NewPieChart()

// 3D Pie
pie3d := ppt.NewPie3DChart()

// Doughnut
doughnut := ppt.NewDoughnutChart()
doughnut.HoleSize = 75 // 10-90

// Scatter
scatter := ppt.NewScatterChart()
scatter.SetSmooth(true)

// Radar
radar := ppt.NewRadarChart()
```

#### Chart Series

```go
s := ppt.NewChartSeriesOrdered("Series Name",
    []string{"Cat1", "Cat2", "Cat3"},
    []float64{10, 20, 30},
)
s.SetFillColor(ppt.ColorRed)
s.SetLabelPosition(ppt.LabelOutsideEnd)
s.ShowValue = true
s.ShowCategoryName = true
s.ShowPercentage = true
s.ShowSeriesName = true
s.Separator = ", "
s.Marker = &ppt.SeriesMarker{Symbol: ppt.MarkerCircle, Size: 5}
```

#### Chart Axes

```go
axX := chart.GetPlotArea().GetAxisX()
axX.SetTitle("Category").SetVisible(true)
axX.SetReversedOrder(true)
axX.SetMajorGridlines(ppt.NewGridlines())

axY := chart.GetPlotArea().GetAxisY()
axY.SetTitle("Value").SetMinBounds(0).SetMaxBounds(100)
axY.SetMajorUnit(20).SetMinorUnit(5)
axY.SetMinorGridlines(&ppt.Gridlines{Width: 1, Color: ppt.ColorBlack})
```

---

### Styles

#### Color

```go
ppt.ColorBlack   // FF000000
ppt.ColorWhite   // FFFFFFFF
ppt.ColorRed     // FFFF0000
ppt.ColorGreen   // FF00FF00
ppt.ColorBlue    // FF0000FF
ppt.ColorYellow  // FFFFFF00

custom := ppt.NewColor("FF8800")     // RGB (auto-adds FF alpha)
custom2 := ppt.NewColor("80FF8800")  // ARGB with transparency
```

#### Font

```go
font := ppt.NewFont()
font.SetName("Arial").SetSize(12)
font.SetBold(true).SetItalic(true)
font.SetColor(ppt.ColorRed)
font.SetUnderline(ppt.UnderlineSingle) // none, sng, dbl, heavy, dash, wavy
font.SetStrikethrough(true)
```

#### Fill

```go
solid := ppt.NewFill().SetSolid(ppt.ColorBlue)
gradient := ppt.NewFill().SetGradientLinear(ppt.ColorRed, ppt.ColorBlue, 90)
```

#### Border

```go
border := &ppt.Border{
    Style: ppt.BorderSolid, // none, solid, dash, dot
    Width: 2,
    Color: ppt.ColorBlack,
}
```

#### Shadow

```go
shadow := ppt.NewShadow()
shadow.SetVisible(true).SetDirection(45).SetDistance(5)
shadow.BlurRadius = 3
shadow.Color = ppt.Color{ARGB: "80000000"}
shadow.Alpha = 50
```

#### Alignment

```go
align := ppt.NewAlignment()
align.SetHorizontal(ppt.HorizontalCenter) // l, ctr, r, just, dist
align.SetVertical(ppt.VerticalMiddle)      // t, ctr, b
align.Level = 2 // indentation level
```

#### Hyperlink

```go
ppt.NewHyperlink("https://example.com")       // external
ppt.NewInternalHyperlink(2)                     // link to slide 2
```

---

### Bullets

```go
// Character bullet
bullet := ppt.NewBullet().SetCharBullet("•", "Arial")
bullet.SetColor(ppt.ColorRed).SetSize(120)

// Numeric bullet
bullet2 := ppt.NewBullet().SetNumericBullet(ppt.NumFormatArabicPeriod, 1)

para.SetBullet(bullet)
```

| Numeric Format | Constant |
|---|---|
| 1. 2. 3. | `NumFormatArabicPeriod` |
| 1) 2) 3) | `NumFormatArabicParen` |
| I. II. III. | `NumFormatRomanUcPeriod` |
| i. ii. iii. | `NumFormatRomanLcPeriod` |
| A. B. C. | `NumFormatAlphaUcPeriod` |
| a. b. c. | `NumFormatAlphaLcPeriod` |

---

### Comments

```go
author := ppt.NewCommentAuthor("John Doe", "JD")
comment := ppt.NewComment()
comment.SetAuthor(author).SetText("Review this").SetPosition(100, 200)
comment.SetDate(time.Now())
slide.AddComment(comment)
```

---

### Writer / Reader

```go
// Write to file
w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
w.(*ppt.PPTXWriter).Save("output.pptx")

// Write to io.Writer
var buf bytes.Buffer
w.WriteTo(&buf)

// Read from file
reader := &ppt.PPTXReader{}
pres, err := reader.Read("input.pptx")

// Read from io.ReaderAt
pres, err := reader.ReadFromReader(readerAt, size)
```

---
---

<a id="中文"></a>

## 中文

包路径：`github.com/VantageDataChat/GoPPT`

所有尺寸使用 EMU（英制公制单位）：1 英寸 = 914400 EMU，1 厘米 = 360000 EMU，1 磅 = 12700 EMU。

---

### 演示文稿 (Presentation)

根对象，代表一个 PowerPoint 文件。

```go
// 创建新演示文稿（自动包含一张空白幻灯片）
p := ppt.New()

// 文档属性
p.GetDocumentProperties().Title = "标题"
p.GetDocumentProperties().Creator = "作者"
p.GetDocumentProperties().Description = "描述"

// 自定义属性
p.GetDocumentProperties().SetCustomProperty("版本", "1.0", ppt.PropertyTypeString)

// 演示文稿属性
p.GetPresentationProperties().SetZoom(1.5)
p.GetPresentationProperties().SetLastView(ppt.ViewSlide)
p.GetPresentationProperties().SetSlideshowType(ppt.SlideshowTypePresent)

// 布局
p.GetLayout().SetLayout(ppt.LayoutScreen16x9)
p.GetLayout().SetCustomLayout(9144000, 6858000) // 自定义 EMU 尺寸
```

| 布局常量 | 说明 |
|---|---|
| `LayoutScreen4x3` | 10" × 7.5"（默认） |
| `LayoutScreen16x9` | 13.33" × 7.5" |
| `LayoutScreen16x10` | 12" × 7.5" |
| `LayoutA4` | A4 横向 |
| `LayoutLetter` | US Letter |
| `LayoutCustom` | 自定义尺寸 |

---

### 幻灯片 (Slide)

```go
slide := p.CreateSlide()           // 创建并添加
slide := p.GetActiveSlide()        // 获取当前活动幻灯片
p.SetActiveSlideIndex(1)           // 切换活动幻灯片
slide, _ := p.GetSlide(0)         // 按索引获取
slides := p.GetAllSlides()         // 获取全部
count := p.GetSlideCount()         // 计数
p.RemoveSlideByIndex(0)            // 删除

slide.SetName("简介")
slide.SetNotes("演讲者备注")
slide.SetVisible(true)
slide.SetBackground(ppt.NewFill().SetSolid(ppt.ColorWhite))
```

---

### 形状 (Shapes)

所有形状共享 `BaseShape`，包含位置、大小、填充、边框、阴影和超链接。

```go
// BaseShape 通用方法
shape.BaseShape.SetOffsetX(914400)   // 距左 1 英寸
shape.BaseShape.SetOffsetY(914400)   // 距顶 1 英寸
shape.BaseShape.SetWidth(5000000)
shape.BaseShape.SetHeight(3000000)
shape.BaseShape.SetName("我的形状")
shape.BaseShape.SetRotation(45)      // 度
shape.BaseShape.SetFill(fill)
shape.BaseShape.SetBorder(border)
shape.BaseShape.SetShadow(shadow)
shape.BaseShape.SetHyperlink(ppt.NewHyperlink("https://example.com"))
```

#### 富文本形状 (RichTextShape)

```go
rt := slide.CreateRichTextShape()
rt.SetOffsetX(100).SetOffsetY(100).SetWidth(8000000).SetHeight(1000000)
rt.SetWordWrap(true)
rt.SetAutoFit(ppt.AutoFitNormal)
rt.SetColumns(2)

// 段落和文本运行
para := rt.GetActiveParagraph()
tr := para.CreateTextRun("你好世界")
tr.GetFont().SetBold(true).SetSize(24).SetColor(ppt.ColorRed).SetName("微软雅黑")

// 换行
para.CreateBreak()
para.CreateTextRun("第二行")

// 新段落
para2 := rt.CreateParagraph()
para2.GetAlignment().SetHorizontal(ppt.HorizontalCenter)
para2.SetLineSpacing(200)
```

#### 图片形状 (DrawingShape)

```go
// 从字节数据
img := slide.CreateDrawingShape()
img.SetImageData(imageBytes, "image/png")
img.SetWidth(2000000).SetHeight(1500000)

// 从文件路径
img2 := ppt.NewDrawingShape()
img2.SetPath("/path/to/image.jpg")
slide.AddShape(img2)
```

支持格式：PNG、JPEG、GIF、BMP、SVG。

#### 表格形状 (TableShape)

```go
table := slide.CreateTableShape(3, 4) // 3 行 4 列
table.SetWidth(8000000).SetHeight(2000000)

cell := table.GetCell(0, 0)
cell.SetText("表头")
cell.SetFill(ppt.NewFill().SetSolid(ppt.ColorBlue))
cell.SetColSpan(2)
```

#### 自动形状 (AutoShape)

```go
shape := slide.CreateAutoShape()
shape.SetAutoShapeType(ppt.AutoShapeEllipse)
shape.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(2000000).SetHeight(1000000)
shape.SetText("形状内文字")
shape.BaseShape.SetFill(ppt.NewFill().SetSolid(ppt.ColorYellow))
```

| 形状类型 | 常量 |
|---|---|
| 矩形 | `AutoShapeRectangle` |
| 圆角矩形 | `AutoShapeRoundedRect` |
| 椭圆 | `AutoShapeEllipse` |
| 三角形 | `AutoShapeTriangle` |
| 菱形 | `AutoShapeDiamond` |
| 五边形 | `AutoShapePentagon` |
| 六边形 | `AutoShapeHexagon` |
| 星形 | `AutoShapeStar4`, `AutoShapeStar5` |
| 箭头 | `AutoShapeArrowRight/Left/Up/Down` |
| 心形 | `AutoShapeHeart` |
| 闪电 | `AutoShapeLightningBolt` |

#### 线条形状 (LineShape)

```go
line := slide.CreateLineShape()
line.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(0)
line.SetLineWidth(2).SetLineColor(ppt.ColorRed)
```

#### 组合形状 (GroupShape)

```go
group := slide.CreateGroupShape()
group.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)

child := ppt.NewRichTextShape()
child.SetOffsetX(0).SetOffsetY(0).SetWidth(2000000).SetHeight(500000)
child.CreateTextRun("组内文字")
group.AddShape(child)
```

#### 占位符形状 (PlaceholderShape)

```go
ph := slide.CreatePlaceholderShape(ppt.PlaceholderTitle)
ph.BaseShape.SetOffsetX(500000).SetOffsetY(300000).SetWidth(8000000).SetHeight(1000000)
ph.CreateTextRun("幻灯片标题")
```

| 占位符类型 | 常量 |
|---|---|
| 标题 | `PlaceholderTitle` |
| 正文 | `PlaceholderBody` |
| 居中标题 | `PlaceholderCtrTitle` |
| 副标题 | `PlaceholderSubTitle` |
| 日期 | `PlaceholderDate` |
| 页脚 | `PlaceholderFooter` |
| 页码 | `PlaceholderSlideNum` |

---

### 图表 (Charts)

```go
chart := slide.CreateChartShape()
chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
chart.BaseShape.SetWidth(7000000).SetHeight(4500000)

// 标题
chart.GetTitle().SetText("我的图表").SetVisible(true)

// 图例
chart.GetLegend().Visible = true
chart.GetLegend().Position = ppt.LegendBottom
```

#### 图表类型

```go
// 柱状图
bar := ppt.NewBarChart()
bar.SetBarGrouping(ppt.BarGroupingClustered) // clustered, stacked, percentStacked
bar.AddSeries(ppt.NewChartSeriesOrdered("销售额", categories, values))

// 3D 柱状图
bar3d := ppt.NewBar3DChart()

// 折线图
line := ppt.NewLineChart()
line.SetSmooth(true)

// 面积图
area := ppt.NewAreaChart()

// 饼图 / 3D 饼图
pie := ppt.NewPieChart()
pie3d := ppt.NewPie3DChart()

// 环形图
doughnut := ppt.NewDoughnutChart()
doughnut.HoleSize = 75

// 散点图
scatter := ppt.NewScatterChart()

// 雷达图
radar := ppt.NewRadarChart()
```

#### 数据系列

```go
s := ppt.NewChartSeriesOrdered("系列名称",
    []string{"类别1", "类别2", "类别3"},
    []float64{10, 20, 30},
)
s.SetFillColor(ppt.ColorRed)
s.ShowValue = true
s.ShowPercentage = true
s.Marker = &ppt.SeriesMarker{Symbol: ppt.MarkerCircle, Size: 5}
```

#### 坐标轴

```go
axX := chart.GetPlotArea().GetAxisX()
axX.SetTitle("类别").SetVisible(true)
axX.SetMajorGridlines(ppt.NewGridlines())

axY := chart.GetPlotArea().GetAxisY()
axY.SetTitle("数值").SetMinBounds(0).SetMaxBounds(100)
axY.SetMajorUnit(20).SetMinorUnit(5)
```

---

### 样式 (Styles)

#### 颜色

```go
ppt.ColorBlack   // FF000000
ppt.ColorWhite   // FFFFFFFF
ppt.ColorRed     // FFFF0000
ppt.ColorGreen   // FF00FF00
ppt.ColorBlue    // FF0000FF
ppt.ColorYellow  // FFFFFF00

custom := ppt.NewColor("FF8800")     // RGB（自动添加 FF 透明度）
custom2 := ppt.NewColor("80FF8800")  // ARGB 含透明度
```

#### 字体

```go
font := ppt.NewFont()
font.SetName("微软雅黑").SetSize(12)
font.SetBold(true).SetItalic(true)
font.SetColor(ppt.ColorRed)
font.SetUnderline(ppt.UnderlineSingle) // none, sng, dbl, heavy, dash, wavy
font.SetStrikethrough(true)
```

#### 填充

```go
solid := ppt.NewFill().SetSolid(ppt.ColorBlue)
gradient := ppt.NewFill().SetGradientLinear(ppt.ColorRed, ppt.ColorBlue, 90)
```

#### 边框

```go
border := &ppt.Border{
    Style: ppt.BorderSolid, // none, solid, dash, dot
    Width: 2,
    Color: ppt.ColorBlack,
}
```

#### 阴影

```go
shadow := ppt.NewShadow()
shadow.SetVisible(true).SetDirection(45).SetDistance(5)
```

#### 对齐

```go
align := ppt.NewAlignment()
align.SetHorizontal(ppt.HorizontalCenter) // l, ctr, r, just, dist
align.SetVertical(ppt.VerticalMiddle)      // t, ctr, b
```

#### 超链接

```go
ppt.NewHyperlink("https://example.com")  // 外部链接
ppt.NewInternalHyperlink(2)               // 链接到第 2 张幻灯片
```

---

### 项目符号 (Bullets)

```go
// 字符符号
bullet := ppt.NewBullet().SetCharBullet("•", "Arial")
bullet.SetColor(ppt.ColorRed).SetSize(120)

// 数字编号
bullet2 := ppt.NewBullet().SetNumericBullet(ppt.NumFormatArabicPeriod, 1)

para.SetBullet(bullet)
```

| 编号格式 | 常量 |
|---|---|
| 1. 2. 3. | `NumFormatArabicPeriod` |
| 1) 2) 3) | `NumFormatArabicParen` |
| I. II. III. | `NumFormatRomanUcPeriod` |
| i. ii. iii. | `NumFormatRomanLcPeriod` |
| A. B. C. | `NumFormatAlphaUcPeriod` |
| a. b. c. | `NumFormatAlphaLcPeriod` |

---

### 批注 (Comments)

```go
author := ppt.NewCommentAuthor("张三", "ZS")
comment := ppt.NewComment()
comment.SetAuthor(author).SetText("请审阅").SetPosition(100, 200)
slide.AddComment(comment)
```

---

### 读写 (Writer / Reader)

```go
// 写入文件
w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
w.(*ppt.PPTXWriter).Save("输出.pptx")

// 写入 io.Writer
var buf bytes.Buffer
w.WriteTo(&buf)

// 从文件读取
reader := &ppt.PPTXReader{}
pres, err := reader.Read("输入.pptx")

// 从 io.ReaderAt 读取
pres, err := reader.ReadFromReader(readerAt, size)
```
