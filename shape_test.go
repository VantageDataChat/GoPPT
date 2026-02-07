package gopresentation

import "testing"

func TestBaseShape(t *testing.T) {
	bs := &BaseShape{}
	bs.SetName("test").SetOffsetX(100).SetOffsetY(200).SetWidth(300).SetHeight(400).SetRotation(45)

	if bs.GetName() != "test" {
		t.Errorf("expected name 'test', got '%s'", bs.GetName())
	}
	if bs.GetOffsetX() != 100 {
		t.Errorf("expected offsetX 100, got %d", bs.GetOffsetX())
	}
	if bs.GetOffsetY() != 200 {
		t.Errorf("expected offsetY 200, got %d", bs.GetOffsetY())
	}
	if bs.GetWidth() != 300 {
		t.Errorf("expected width 300, got %d", bs.GetWidth())
	}
	if bs.GetHeight() != 400 {
		t.Errorf("expected height 400, got %d", bs.GetHeight())
	}
	if bs.GetRotation() != 45 {
		t.Errorf("expected rotation 45, got %d", bs.GetRotation())
	}

	bs.SetDescription("desc")
	if bs.GetDescription() != "desc" {
		t.Errorf("expected description 'desc', got '%s'", bs.GetDescription())
	}

	// Fill
	fill := bs.GetFill()
	if fill == nil {
		t.Fatal("GetFill() returned nil")
	}
	if fill.Type != FillNone {
		t.Error("expected FillNone by default")
	}

	// Border
	border := bs.GetBorder()
	if border == nil {
		t.Fatal("GetBorder() returned nil")
	}

	// Shadow
	shadow := bs.GetShadow()
	if shadow == nil {
		t.Fatal("GetShadow() returned nil")
	}

	// Hyperlink
	if bs.GetHyperlink() != nil {
		t.Error("hyperlink should be nil by default")
	}
	hl := NewHyperlink("https://example.com")
	bs.SetHyperlink(hl)
	if bs.GetHyperlink().URL != "https://example.com" {
		t.Error("hyperlink URL mismatch")
	}
}

func TestRichTextShape(t *testing.T) {
	rt := NewRichTextShape()

	rt.SetHeight(300).SetWidth(600).SetOffsetX(100).SetOffsetY(200)
	if rt.GetHeight() != 300 {
		t.Errorf("expected height 300, got %d", rt.GetHeight())
	}
	if rt.GetWidth() != 600 {
		t.Errorf("expected width 600, got %d", rt.GetWidth())
	}

	// Active paragraph
	para := rt.GetActiveParagraph()
	if para == nil {
		t.Fatal("GetActiveParagraph() returned nil")
	}

	// Create text run
	tr := rt.CreateTextRun("Hello World")
	if tr.GetText() != "Hello World" {
		t.Errorf("expected 'Hello World', got '%s'", tr.GetText())
	}

	// Create break
	br := rt.CreateBreak()
	if br == nil {
		t.Fatal("CreateBreak() returned nil")
	}

	// Create new paragraph
	p2 := rt.CreateParagraph()
	if p2 == nil {
		t.Fatal("CreateParagraph() returned nil")
	}
	if len(rt.GetParagraphs()) != 2 {
		t.Errorf("expected 2 paragraphs, got %d", len(rt.GetParagraphs()))
	}

	// Auto fit
	rt.SetAutoFit(AutoFitNormal)
	if rt.GetAutoFit() != AutoFitNormal {
		t.Error("expected AutoFitNormal")
	}

	// Word wrap
	if !rt.GetWordWrap() {
		t.Error("word wrap should be true by default")
	}
	rt.SetWordWrap(false)
	if rt.GetWordWrap() {
		t.Error("word wrap should be false")
	}

	// Columns
	if rt.GetColumns() != 1 {
		t.Errorf("expected 1 column, got %d", rt.GetColumns())
	}
	rt.SetColumns(2)
	if rt.GetColumns() != 2 {
		t.Errorf("expected 2 columns, got %d", rt.GetColumns())
	}
}

func TestParagraph(t *testing.T) {
	p := NewParagraph()

	// Alignment
	align := p.GetAlignment()
	if align == nil {
		t.Fatal("GetAlignment() returned nil")
	}
	if align.Horizontal != HorizontalLeft {
		t.Errorf("expected HorizontalLeft, got '%s'", align.Horizontal)
	}

	// Line spacing
	p.SetLineSpacing(200)
	if p.GetLineSpacing() != 200 {
		t.Errorf("expected line spacing 200, got %d", p.GetLineSpacing())
	}

	// Create text run
	tr := p.CreateTextRun("Test")
	if tr == nil {
		t.Fatal("CreateTextRun() returned nil")
	}
	if tr.GetElementType() != "textrun" {
		t.Errorf("expected 'textrun', got '%s'", tr.GetElementType())
	}

	// Create break
	br := p.CreateBreak()
	if br.GetElementType() != "break" {
		t.Errorf("expected 'break', got '%s'", br.GetElementType())
	}

	elements := p.GetElements()
	if len(elements) != 2 {
		t.Errorf("expected 2 elements, got %d", len(elements))
	}
}

func TestTextRun(t *testing.T) {
	tr := &TextRun{text: "Hello", font: NewFont()}

	tr.SetText("World")
	if tr.GetText() != "World" {
		t.Errorf("expected 'World', got '%s'", tr.GetText())
	}

	font := tr.GetFont()
	if font == nil {
		t.Fatal("GetFont() returned nil")
	}

	newFont := NewFont()
	newFont.SetBold(true)
	tr.SetFont(newFont)
	if !tr.GetFont().Bold {
		t.Error("font should be bold")
	}

	// Hyperlink
	if tr.GetHyperlink() != nil {
		t.Error("hyperlink should be nil by default")
	}
	hl := NewHyperlink("https://example.com")
	tr.SetHyperlink(hl)
	if tr.GetHyperlink().URL != "https://example.com" {
		t.Error("hyperlink URL mismatch")
	}
}

func TestDrawingShape(t *testing.T) {
	ds := NewDrawingShape()

	ds.SetPath("/path/to/image.png")
	if ds.GetPath() != "/path/to/image.png" {
		t.Errorf("expected path, got '%s'", ds.GetPath())
	}

	ds.SetImageData([]byte{1, 2, 3}, "image/png")
	if len(ds.GetImageData()) != 3 {
		t.Errorf("expected 3 bytes, got %d", len(ds.GetImageData()))
	}
	if ds.GetMimeType() != "image/png" {
		t.Errorf("expected 'image/png', got '%s'", ds.GetMimeType())
	}

	ds.SetHeight(100).SetWidth(200).SetOffsetX(10).SetOffsetY(20)
	if ds.GetHeight() != 100 || ds.GetWidth() != 200 {
		t.Error("dimensions mismatch")
	}
}

func TestAutoShape(t *testing.T) {
	as := NewAutoShape()

	if as.GetAutoShapeType() != AutoShapeRectangle {
		t.Error("expected default rectangle")
	}

	as.SetAutoShapeType(AutoShapeEllipse)
	if as.GetAutoShapeType() != AutoShapeEllipse {
		t.Error("expected ellipse")
	}

	as.SetText("Hello")
	if as.GetText() != "Hello" {
		t.Errorf("expected 'Hello', got '%s'", as.GetText())
	}
}

func TestLineShape(t *testing.T) {
	ls := NewLineShape()

	if ls.GetLineStyle() != BorderSolid {
		t.Error("expected solid line style")
	}

	ls.SetLineStyle(BorderDash)
	if ls.GetLineStyle() != BorderDash {
		t.Error("expected dash line style")
	}

	ls.SetLineWidth(3)
	if ls.GetLineWidth() != 3 {
		t.Errorf("expected width 3, got %d", ls.GetLineWidth())
	}

	ls.SetLineColor(ColorRed)
	if ls.GetLineColor() != ColorRed {
		t.Error("expected red color")
	}
}

func TestTableShape(t *testing.T) {
	ts := NewTableShape(3, 4)

	if ts.GetNumRows() != 3 {
		t.Errorf("expected 3 rows, got %d", ts.GetNumRows())
	}
	if ts.GetNumCols() != 4 {
		t.Errorf("expected 4 cols, got %d", ts.GetNumCols())
	}

	// Get cell
	cell := ts.GetCell(0, 0)
	if cell == nil {
		t.Fatal("GetCell(0,0) returned nil")
	}

	// Out of range
	if ts.GetCell(10, 10) != nil {
		t.Error("expected nil for out of range cell")
	}

	// Set text
	cell.SetText("Hello")
	paras := cell.GetParagraphs()
	if len(paras) != 1 {
		t.Errorf("expected 1 paragraph, got %d", len(paras))
	}

	// Fill
	cell.SetFill(NewFill().SetSolid(ColorBlue))
	if cell.GetFill().Type != FillSolid {
		t.Error("expected solid fill")
	}

	// Borders
	borders := cell.GetBorders()
	if borders == nil {
		t.Fatal("GetBorders() returned nil")
	}

	// Span
	cell.SetColSpan(2)
	if cell.GetColSpan() != 2 {
		t.Errorf("expected col span 2, got %d", cell.GetColSpan())
	}
	cell.SetRowSpan(3)
	if cell.GetRowSpan() != 3 {
		t.Errorf("expected row span 3, got %d", cell.GetRowSpan())
	}

	// Rows
	rows := ts.GetRows()
	if len(rows) != 3 {
		t.Errorf("expected 3 rows, got %d", len(rows))
	}

	// Chaining
	ts.SetHeight(1000).SetWidth(2000)
	if ts.GetHeight() != 1000 || ts.GetWidth() != 2000 {
		t.Error("dimensions mismatch")
	}
}
