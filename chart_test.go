package gopresentation

import (
	"archive/zip"
	"bytes"
	"strings"
	"testing"
)

func TestBarChart(t *testing.T) {
	bc := NewBarChart()
	if bc.GetChartTypeName() != "bar" {
		t.Error("expected bar chart type name")
	}

	s := NewChartSeriesOrdered("Sales", []string{"Q1", "Q2", "Q3", "Q4"}, []float64{100, 200, 150, 300})
	bc.AddSeries(s)
	if len(bc.Series) != 1 {
		t.Errorf("expected 1 series, got %d", len(bc.Series))
	}

	bc.SetBarGrouping(BarGroupingStacked)
	if bc.BarGrouping != BarGroupingStacked {
		t.Error("expected stacked grouping")
	}
	if bc.OverlapPercent != 100 {
		t.Errorf("expected overlap 100 for stacked, got %d", bc.OverlapPercent)
	}

	bc.SetGapWidthPercent(250)
	if bc.GapWidthPercent != 250 {
		t.Errorf("expected gap 250, got %d", bc.GapWidthPercent)
	}

	bc.SetGapWidthPercent(600) // should clamp to 500
	if bc.GapWidthPercent != 500 {
		t.Errorf("expected gap clamped to 500, got %d", bc.GapWidthPercent)
	}

	bc.SetOverlapPercent(-50)
	if bc.OverlapPercent != -50 {
		t.Errorf("expected overlap -50, got %d", bc.OverlapPercent)
	}
}

func TestBar3DChart(t *testing.T) {
	bc := NewBar3DChart()
	if bc.GetChartTypeName() != "bar3D" {
		t.Error("expected bar3D chart type name")
	}
}

func TestLineChart(t *testing.T) {
	lc := NewLineChart()
	if lc.GetChartTypeName() != "line" {
		t.Error("expected line chart type name")
	}

	s := NewChartSeriesOrdered("Trend", []string{"Jan", "Feb", "Mar"}, []float64{10, 20, 15})
	lc.AddSeries(s)
	lc.SetSmooth(true)
	if !lc.IsSmooth {
		t.Error("expected smooth line")
	}
}

func TestPieChart(t *testing.T) {
	pc := NewPieChart()
	if pc.GetChartTypeName() != "pie" {
		t.Error("expected pie chart type name")
	}
	s := NewChartSeries("Market Share", map[string]float64{"A": 40, "B": 30, "C": 30})
	pc.AddSeries(s)
	if len(pc.Series) != 1 {
		t.Errorf("expected 1 series, got %d", len(pc.Series))
	}
}

func TestPie3DChart(t *testing.T) {
	pc := NewPie3DChart()
	if pc.GetChartTypeName() != "pie3D" {
		t.Error("expected pie3D chart type name")
	}
}

func TestDoughnutChart(t *testing.T) {
	dc := NewDoughnutChart()
	if dc.GetChartTypeName() != "doughnut" {
		t.Error("expected doughnut chart type name")
	}
	if dc.HoleSize != 50 {
		t.Errorf("expected hole size 50, got %d", dc.HoleSize)
	}
	s := NewChartSeriesOrdered("Data", []string{"A", "B"}, []float64{60, 40})
	dc.AddSeries(s)
}

func TestAreaChart(t *testing.T) {
	ac := NewAreaChart()
	if ac.GetChartTypeName() != "area" {
		t.Error("expected area chart type name")
	}
	s := NewChartSeriesOrdered("Area", []string{"X", "Y"}, []float64{5, 10})
	ac.AddSeries(s)
}

func TestScatterChart(t *testing.T) {
	sc := NewScatterChart()
	if sc.GetChartTypeName() != "scatter" {
		t.Error("expected scatter chart type name")
	}
	s := NewChartSeriesOrdered("Points", []string{"1", "2", "3"}, []float64{1, 4, 9})
	sc.AddSeries(s)
	sc.SetSmooth(true)
	if !sc.IsSmooth {
		t.Error("expected smooth scatter")
	}
}

func TestRadarChart(t *testing.T) {
	rc := NewRadarChart()
	if rc.GetChartTypeName() != "radar" {
		t.Error("expected radar chart type name")
	}
	s := NewChartSeriesOrdered("Skills", []string{"A", "B", "C"}, []float64{8, 6, 9})
	rc.AddSeries(s)
}

func TestChartShape(t *testing.T) {
	cs := NewChartShape()
	if cs.GetType() != ShapeTypeChart {
		t.Error("expected ShapeTypeChart")
	}

	cs.GetTitle().SetText("My Chart").SetVisible(true)
	if cs.GetTitle().Text != "My Chart" {
		t.Error("title text mismatch")
	}

	cs.SetDisplayBlankAs(ChartBlankAsGap)
	if cs.GetDisplayBlankAs() != ChartBlankAsGap {
		t.Error("display blank as mismatch")
	}

	cs.GetLegend().Visible = true
	cs.GetLegend().Position = LegendRight
	if cs.GetLegend().Position != LegendRight {
		t.Error("legend position mismatch")
	}
}

func TestChartAxis(t *testing.T) {
	ax := NewChartAxis()
	ax.SetTitle("X Axis").SetTitleRotation(45).SetVisible(true)
	if ax.Title != "X Axis" {
		t.Error("title mismatch")
	}

	ax.SetMinBounds(0).SetMaxBounds(100)
	if *ax.MinBounds != 0 || *ax.MaxBounds != 100 {
		t.Error("bounds mismatch")
	}
	ax.ClearMinBounds().ClearMaxBounds()
	if ax.MinBounds != nil || ax.MaxBounds != nil {
		t.Error("bounds should be nil after clear")
	}

	ax.SetMajorUnit(10).SetMinorUnit(2)
	if *ax.MajorUnit != 10 || *ax.MinorUnit != 2 {
		t.Error("unit mismatch")
	}

	ax.SetCrossesAt(AxisCrossesMax)
	if ax.CrossesAt != AxisCrossesMax {
		t.Error("crosses at mismatch")
	}

	ax.SetReversedOrder(true)
	if !ax.ReversedOrder {
		t.Error("expected reversed order")
	}

	gl := NewGridlines()
	ax.SetMajorGridlines(gl).SetMinorGridlines(gl)
	if ax.MajorGridlines == nil || ax.MinorGridlines == nil {
		t.Error("gridlines should not be nil")
	}

	ax.SetMajorTickMark(TickMarkInside).SetMinorTickMark(TickMarkOutside)
	if ax.MajorTickMark != TickMarkInside {
		t.Error("major tick mark mismatch")
	}

	ax.SetTickLabelPosition(TickLabelPosLow)
	if ax.TickLabelPos != TickLabelPosLow {
		t.Error("tick label position mismatch")
	}
}

func TestChartSeries(t *testing.T) {
	s := NewChartSeriesOrdered("Sales", []string{"Q1", "Q2"}, []float64{100, 200})
	if s.Title != "Sales" {
		t.Error("title mismatch")
	}
	if len(s.Categories) != 2 {
		t.Errorf("expected 2 categories, got %d", len(s.Categories))
	}

	s.SetFillColor(ColorBlue).SetLabelPosition(LabelInsideEnd)
	if s.FillColor != ColorBlue {
		t.Error("fill color mismatch")
	}
	if s.LabelPosition != LabelInsideEnd {
		t.Error("label position mismatch")
	}
}

func TestView3D(t *testing.T) {
	v := NewView3D()
	if v.RotX != 15 || v.RotY != 20 {
		t.Error("default rotation mismatch")
	}
	v.SetHeightPercent(nil)
	if v.HeightPercent != nil {
		t.Error("height percent should be nil for autoscale")
	}
}

func TestWriteBarChartSlide(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000).SetWidth(8000000).SetHeight(5000000)
	chart.GetTitle().SetText("Sales Report")

	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("2024", []string{"Q1", "Q2", "Q3", "Q4"}, []float64{100, 200, 150, 300}))
	bar.AddSeries(NewChartSeriesOrdered("2025", []string{"Q1", "Q2", "Q3", "Q4"}, []float64{120, 180, 220, 280}))
	chart.GetPlotArea().SetType(bar)
	chart.GetPlotArea().GetAxisX().SetTitle("Quarter")
	chart.GetPlotArea().GetAxisY().SetTitle("Revenue")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Check chart file exists
	chartData, err := readFileFromZip(zr, "ppt/charts/chart1.xml")
	if err != nil {
		t.Fatal("chart1.xml not found in zip")
	}
	content := string(chartData)
	if !strings.Contains(content, "barChart") {
		t.Error("chart should contain barChart element")
	}
	if !strings.Contains(content, "Sales Report") {
		t.Error("chart should contain title")
	}
	if !strings.Contains(content, "Q1") {
		t.Error("chart should contain category data")
	}

	// Check content types
	ctData, _ := readFileFromZip(zr, "[Content_Types].xml")
	if !strings.Contains(string(ctData), "chart") {
		t.Error("content types should include chart")
	}

	// Check slide references chart
	slideData, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	if !strings.Contains(string(slideData), "graphicFrame") {
		t.Error("slide should contain graphicFrame for chart")
	}
}

func TestWriteLineChartSlide(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(5000000).SetHeight(3000000)

	lc := NewLineChart()
	lc.AddSeries(NewChartSeriesOrdered("Trend", []string{"Jan", "Feb", "Mar"}, []float64{10, 20, 15}))
	lc.SetSmooth(true)
	chart.GetPlotArea().SetType(lc)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	chartData, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(chartData)
	if !strings.Contains(content, "lineChart") {
		t.Error("chart should contain lineChart element")
	}
	if !strings.Contains(content, `smooth val="1"`) {
		t.Error("chart should have smooth enabled")
	}
}

func TestWritePieChartSlide(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(5000000).SetHeight(3000000)

	pc := NewPieChart()
	s := NewChartSeriesOrdered("Share", []string{"A", "B", "C"}, []float64{40, 35, 25})
	s.ShowPercentage = true
	s.ShowValue = true
	pc.AddSeries(s)
	chart.GetPlotArea().SetType(pc)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	chartData, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(chartData)
	if !strings.Contains(content, "pieChart") {
		t.Error("chart should contain pieChart element")
	}
	if !strings.Contains(content, "showPercent") {
		t.Error("chart should show percentage")
	}
}

func TestWriteScatterChartSlide(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(5000000).SetHeight(3000000)

	sc := NewScatterChart()
	sc.AddSeries(NewChartSeriesOrdered("Points", []string{"1", "2", "3"}, []float64{1, 4, 9}))
	chart.GetPlotArea().SetType(sc)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	chartData, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	if !strings.Contains(string(chartData), "scatterChart") {
		t.Error("chart should contain scatterChart element")
	}
}

func TestWriteDoughnutChartSlide(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(5000000).SetHeight(3000000)

	dc := NewDoughnutChart()
	dc.HoleSize = 75
	dc.AddSeries(NewChartSeriesOrdered("Data", []string{"X", "Y"}, []float64{60, 40}))
	chart.GetPlotArea().SetType(dc)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	chartData, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(chartData)
	if !strings.Contains(content, "doughnutChart") {
		t.Error("chart should contain doughnutChart element")
	}
	if !strings.Contains(content, `holeSize val="75"`) {
		t.Error("chart should have hole size 75")
	}
}

func TestWriteRadarChartSlide(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(5000000).SetHeight(3000000)

	rc := NewRadarChart()
	rc.AddSeries(NewChartSeriesOrdered("Skills", []string{"A", "B", "C", "D"}, []float64{8, 6, 9, 7}))
	chart.GetPlotArea().SetType(rc)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	chartData, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	if !strings.Contains(string(chartData), "radarChart") {
		t.Error("chart should contain radarChart element")
	}
}

func TestWriteAreaChartSlide(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(5000000).SetHeight(3000000)

	ac := NewAreaChart()
	ac.AddSeries(NewChartSeriesOrdered("Area", []string{"X", "Y", "Z"}, []float64{5, 10, 7}))
	chart.GetPlotArea().SetType(ac)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	chartData, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	if !strings.Contains(string(chartData), "areaChart") {
		t.Error("chart should contain areaChart element")
	}
}

func TestWriteChartWithAxisConfig(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(5000000).SetHeight(3000000)

	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("Data", []string{"A", "B"}, []float64{10, 20}))
	chart.GetPlotArea().SetType(bar)

	chart.GetPlotArea().GetAxisX().SetTitle("Categories")
	chart.GetPlotArea().GetAxisY().SetTitle("Values").SetMinBounds(0).SetMaxBounds(100).SetMajorUnit(10)
	chart.GetPlotArea().GetAxisY().SetMajorGridlines(NewGridlines())

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	chartData, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(chartData)

	if !strings.Contains(content, "Categories") {
		t.Error("chart should contain X axis title")
	}
	if !strings.Contains(content, "Values") {
		t.Error("chart should contain Y axis title")
	}
	if !strings.Contains(content, "majorGridlines") {
		t.Error("chart should contain major gridlines")
	}
}
