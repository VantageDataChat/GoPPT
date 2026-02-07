package gopresentation

import (
	"bytes"
	"os"
	"testing"
)

func TestReaderCreation(t *testing.T) {
	r, err := NewReader(ReaderPowerPoint2007)
	if err != nil {
		t.Fatalf("NewReader error: %v", err)
	}
	if r == nil {
		t.Fatal("reader is nil")
	}

	_, err = NewReader("invalid")
	if err == nil {
		t.Error("expected error for invalid reader type")
	}
}

func TestRoundTripMinimal(t *testing.T) {
	// Create presentation
	p := New()
	props := p.GetDocumentProperties()
	props.Creator = "Round Trip Test"
	props.Title = "Test Presentation"
	props.Description = "Testing round trip"

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	// Verify
	if pres.GetSlideCount() != 1 {
		t.Errorf("expected 1 slide, got %d", pres.GetSlideCount())
	}
	if pres.GetDocumentProperties().Creator != "Round Trip Test" {
		t.Errorf("expected creator 'Round Trip Test', got '%s'", pres.GetDocumentProperties().Creator)
	}
	if pres.GetDocumentProperties().Title != "Test Presentation" {
		t.Errorf("expected title 'Test Presentation', got '%s'", pres.GetDocumentProperties().Title)
	}
}

func TestRoundTripWithText(t *testing.T) {
	// Create presentation with text
	p := New()
	slide := p.GetActiveSlide()
	shape := slide.CreateRichTextShape()
	shape.SetHeight(300).SetWidth(600).SetOffsetX(100).SetOffsetY(200)
	tr := shape.CreateTextRun("Hello World")
	tr.GetFont().SetBold(true).SetSize(24).SetColor(NewColor("FFE06B20"))

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	// Verify slide count
	if pres.GetSlideCount() != 1 {
		t.Fatalf("expected 1 slide, got %d", pres.GetSlideCount())
	}

	// Verify shapes
	shapes := pres.GetActiveSlide().GetShapes()
	if len(shapes) == 0 {
		t.Fatal("expected at least 1 shape")
	}

	// Find the rich text shape
	var rtShape *RichTextShape
	for _, s := range shapes {
		if rt, ok := s.(*RichTextShape); ok {
			rtShape = rt
			break
		}
	}
	if rtShape == nil {
		t.Fatal("expected a RichTextShape")
	}

	// Verify position
	if rtShape.GetOffsetX() != 100 {
		t.Errorf("expected offsetX 100, got %d", rtShape.GetOffsetX())
	}
	if rtShape.GetOffsetY() != 200 {
		t.Errorf("expected offsetY 200, got %d", rtShape.GetOffsetY())
	}
	if rtShape.GetWidth() != 600 {
		t.Errorf("expected width 600, got %d", rtShape.GetWidth())
	}
	if rtShape.GetHeight() != 300 {
		t.Errorf("expected height 300, got %d", rtShape.GetHeight())
	}

	// Verify text content
	paras := rtShape.GetParagraphs()
	if len(paras) == 0 {
		t.Fatal("expected at least 1 paragraph")
	}

	elements := paras[0].GetElements()
	if len(elements) == 0 {
		t.Fatal("expected at least 1 element")
	}

	textRun, ok := elements[0].(*TextRun)
	if !ok {
		t.Fatal("expected TextRun element")
	}
	if textRun.GetText() != "Hello World" {
		t.Errorf("expected 'Hello World', got '%s'", textRun.GetText())
	}

	// Verify font
	font := textRun.GetFont()
	if !font.Bold {
		t.Error("expected bold font")
	}
	if font.Size != 24 {
		t.Errorf("expected font size 24, got %d", font.Size)
	}
}

func TestRoundTripMultipleSlides(t *testing.T) {
	p := New()
	p.GetActiveSlide().SetName("Slide 1")

	s2 := p.CreateSlide()
	s2.SetName("Slide 2")
	rt := s2.CreateRichTextShape()
	rt.SetHeight(100).SetWidth(200)
	rt.CreateTextRun("Slide 2 content")

	s3 := p.CreateSlide()
	s3.SetName("Slide 3")

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	if pres.GetSlideCount() != 3 {
		t.Errorf("expected 3 slides, got %d", pres.GetSlideCount())
	}
}

func TestRoundTripWithImage(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	pngData := createMinimalPNG()
	img := slide.CreateDrawingShape()
	img.SetImageData(pngData, "image/png")
	img.SetHeight(100).SetWidth(100).SetOffsetX(50).SetOffsetY(50)
	img.BaseShape.SetName("Test Image")

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	shapes := pres.GetActiveSlide().GetShapes()
	foundImage := false
	for _, s := range shapes {
		if ds, ok := s.(*DrawingShape); ok {
			foundImage = true
			if len(ds.GetImageData()) == 0 {
				t.Error("image data should not be empty")
			}
			if ds.GetOffsetX() != 50 {
				t.Errorf("expected offsetX 50, got %d", ds.GetOffsetX())
			}
		}
	}
	if !foundImage {
		t.Error("expected to find an image shape")
	}
}

func TestRoundTripWithLine(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	line := slide.CreateLineShape()
	line.BaseShape.SetOffsetX(0).SetOffsetY(1000).SetWidth(5000000).SetHeight(0)
	line.SetLineWidth(2).SetLineColor(ColorRed)

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	shapes := pres.GetActiveSlide().GetShapes()
	foundLine := false
	for _, s := range shapes {
		if ls, ok := s.(*LineShape); ok {
			foundLine = true
			if ls.GetOffsetY() != 1000 {
				t.Errorf("expected offsetY 1000, got %d", ls.GetOffsetY())
			}
		}
	}
	if !foundLine {
		t.Error("expected to find a line shape")
	}
}

func TestRoundTripDocumentProperties(t *testing.T) {
	p := New()
	props := p.GetDocumentProperties()
	props.Creator = "Author Name"
	props.Title = "My Title"
	props.Description = "My Description"
	props.Subject = "My Subject"
	props.Keywords = "key1, key2"
	props.Category = "My Category"
	props.Revision = "1.0"

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	rProps := pres.GetDocumentProperties()
	if rProps.Creator != "Author Name" {
		t.Errorf("expected creator 'Author Name', got '%s'", rProps.Creator)
	}
	if rProps.Title != "My Title" {
		t.Errorf("expected title 'My Title', got '%s'", rProps.Title)
	}
	if rProps.Description != "My Description" {
		t.Errorf("expected description 'My Description', got '%s'", rProps.Description)
	}
	if rProps.Subject != "My Subject" {
		t.Errorf("expected subject 'My Subject', got '%s'", rProps.Subject)
	}
	if rProps.Keywords != "key1, key2" {
		t.Errorf("expected keywords 'key1, key2', got '%s'", rProps.Keywords)
	}
	if rProps.Category != "My Category" {
		t.Errorf("expected category 'My Category', got '%s'", rProps.Category)
	}
	if rProps.Revision != "1.0" {
		t.Errorf("expected revision '1.0', got '%s'", rProps.Revision)
	}
}

func TestRoundTripLayout(t *testing.T) {
	p := New()
	p.GetLayout().SetLayout(LayoutScreen16x9)

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	if pres.GetLayout().CX != 12192000 {
		t.Errorf("expected CX 12192000, got %d", pres.GetLayout().CX)
	}
}

func TestReadFromFile(t *testing.T) {
	// Create a temp file
	p := New()
	p.GetDocumentProperties().Title = "File Read Test"
	slide := p.GetActiveSlide()
	rt := slide.CreateRichTextShape()
	rt.SetHeight(100).SetWidth(200)
	rt.CreateTextRun("File test")

	tmpFile := "test_read.pptx"
	defer os.Remove(tmpFile)

	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.Save(tmpFile)

	// Read from file
	reader := &PPTXReader{}
	pres, err := reader.Read(tmpFile)
	if err != nil {
		t.Fatalf("Read error: %v", err)
	}

	if pres.GetDocumentProperties().Title != "File Read Test" {
		t.Errorf("expected title 'File Read Test', got '%s'", pres.GetDocumentProperties().Title)
	}
	if pres.GetSlideCount() != 1 {
		t.Errorf("expected 1 slide, got %d", pres.GetSlideCount())
	}
}

func TestReadNonExistentFile(t *testing.T) {
	reader := &PPTXReader{}
	_, err := reader.Read("nonexistent.pptx")
	if err == nil {
		t.Error("expected error for non-existent file")
	}
}

func TestReadInvalidZip(t *testing.T) {
	reader := &PPTXReader{}
	data := []byte("not a zip file")
	_, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err == nil {
		t.Error("expected error for invalid zip")
	}
}

func TestRoundTripComplexPresentation(t *testing.T) {
	p := New()
	p.GetDocumentProperties().Creator = "Complex Test"
	p.GetDocumentProperties().Title = "Complex Presentation"
	p.GetLayout().SetLayout(LayoutScreen16x9)

	// Slide 1: Rich text with formatting
	slide1 := p.GetActiveSlide()
	rt1 := slide1.CreateRichTextShape()
	rt1.SetHeight(500).SetWidth(800).SetOffsetX(100).SetOffsetY(100)
	tr1 := rt1.CreateTextRun("Title Text")
	tr1.GetFont().SetBold(true).SetSize(36).SetColor(NewColor("FF0000"))
	rt1.CreateParagraph()
	tr2 := rt1.CreateTextRun("Subtitle")
	tr2.GetFont().SetItalic(true).SetSize(18)

	// Slide 2: Auto shape
	slide2 := p.CreateSlide()
	as := slide2.CreateAutoShape()
	as.SetAutoShapeType(AutoShapeEllipse)
	as.BaseShape.SetOffsetX(200).SetOffsetY(200).SetWidth(400).SetHeight(400)
	as.GetFill().SetSolid(ColorBlue)

	// Slide 3: Image
	slide3 := p.CreateSlide()
	img := slide3.CreateDrawingShape()
	img.SetImageData(createMinimalPNG(), "image/png")
	img.SetHeight(200).SetWidth(200).SetOffsetX(300).SetOffsetY(300)

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	// Verify
	if pres.GetSlideCount() != 3 {
		t.Errorf("expected 3 slides, got %d", pres.GetSlideCount())
	}
	if pres.GetDocumentProperties().Creator != "Complex Test" {
		t.Errorf("expected creator 'Complex Test', got '%s'", pres.GetDocumentProperties().Creator)
	}
	if pres.GetLayout().CX != 12192000 {
		t.Errorf("expected 16:9 layout width, got %d", pres.GetLayout().CX)
	}
}

func TestResolveRelativePath(t *testing.T) {
	tests := []struct {
		base     string
		rel      string
		expected string
	}{
		{"ppt/slides", "../media/image1.png", "ppt/media/image1.png"},
		{"ppt/slides", "image.png", "ppt/slides/image.png"},
		{"ppt", "slides/slide1.xml", "ppt/slides/slide1.xml"},
		{"", "/ppt/slides/slide1.xml", "ppt/slides/slide1.xml"},
	}

	for _, tt := range tests {
		result := resolveRelativePath(tt.base, tt.rel)
		if result != tt.expected {
			t.Errorf("resolveRelativePath(%q, %q) = %q, want %q", tt.base, tt.rel, result, tt.expected)
		}
	}
}

func TestGuessMimeType(t *testing.T) {
	tests := []struct {
		path     string
		expected string
	}{
		{"image.png", "image/png"},
		{"image.PNG", "image/png"},
		{"photo.jpg", "image/jpeg"},
		{"photo.jpeg", "image/jpeg"},
		{"anim.gif", "image/gif"},
		{"icon.bmp", "image/bmp"},
		{"vector.svg", "image/svg+xml"},
		{"unknown.xyz", "image/png"},
	}

	for _, tt := range tests {
		result := guessMimeType(tt.path)
		if result != tt.expected {
			t.Errorf("guessMimeType(%q) = %q, want %q", tt.path, result, tt.expected)
		}
	}
}
