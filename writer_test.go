package gopresentation

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"os"
	"strings"
	"testing"
)

func TestWriterCreation(t *testing.T) {
	p := New()

	// Valid writer
	w, err := NewWriter(p, WriterPowerPoint2007)
	if err != nil {
		t.Fatalf("NewWriter error: %v", err)
	}
	if w == nil {
		t.Fatal("writer is nil")
	}

	// Invalid writer
	_, err = NewWriter(p, "invalid")
	if err == nil {
		t.Error("expected error for invalid writer type")
	}
}

func TestWriteMinimalPresentation(t *testing.T) {
	p := New()
	w, err := NewWriter(p, WriterPowerPoint2007)
	if err != nil {
		t.Fatalf("NewWriter error: %v", err)
	}

	var buf bytes.Buffer
	err = w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	// Verify it's a valid zip
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("not a valid zip: %v", err)
	}

	// Check required files exist
	requiredFiles := []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"docProps/app.xml",
		"docProps/core.xml",
		"ppt/presentation.xml",
		"ppt/_rels/presentation.xml.rels",
		"ppt/presProps.xml",
		"ppt/viewProps.xml",
		"ppt/tableStyles.xml",
		"ppt/slideMasters/slideMaster1.xml",
		"ppt/slideMasters/_rels/slideMaster1.xml.rels",
		"ppt/slideLayouts/slideLayout1.xml",
		"ppt/slideLayouts/_rels/slideLayout1.xml.rels",
		"ppt/theme/theme1.xml",
		"ppt/slides/slide1.xml",
		"ppt/slides/_rels/slide1.xml.rels",
	}

	fileMap := make(map[string]bool)
	for _, f := range zr.File {
		fileMap[f.Name] = true
	}

	for _, req := range requiredFiles {
		if !fileMap[req] {
			t.Errorf("missing required file: %s", req)
		}
	}
}

func TestWriteContentTypesXML(t *testing.T) {
	p := New()
	p.CreateSlide() // 2 slides total

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, err := readFileFromZip(zr, "[Content_Types].xml")
	if err != nil {
		t.Fatalf("failed to read content types: %v", err)
	}

	content := string(data)
	if !strings.Contains(content, "slide1.xml") {
		t.Error("content types should reference slide1")
	}
	if !strings.Contains(content, "slide2.xml") {
		t.Error("content types should reference slide2")
	}
}

func TestWritePresentationXML(t *testing.T) {
	p := New()
	p.GetLayout().SetLayout(LayoutScreen16x9)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, err := readFileFromZip(zr, "ppt/presentation.xml")
	if err != nil {
		t.Fatalf("failed to read presentation.xml: %v", err)
	}

	content := string(data)
	if !strings.Contains(content, `cx="12192000"`) {
		t.Error("presentation should have 16:9 width")
	}
	if !strings.Contains(content, "sldMasterId") {
		t.Error("presentation should reference slide master")
	}
	if !strings.Contains(content, "sldId") {
		t.Error("presentation should reference slides")
	}
}

func TestWriteCoreProperties(t *testing.T) {
	p := New()
	props := p.GetDocumentProperties()
	props.Creator = "Test Author"
	props.Title = "Test Title"
	props.Description = "Test Description"
	props.Company = "Test Company"

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, err := readFileFromZip(zr, "docProps/core.xml")
	if err != nil {
		t.Fatalf("failed to read core.xml: %v", err)
	}

	content := string(data)
	if !strings.Contains(content, "Test Author") {
		t.Error("core properties should contain creator")
	}
	if !strings.Contains(content, "Test Title") {
		t.Error("core properties should contain title")
	}
}

func TestWriteRichTextSlide(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	shape := slide.CreateRichTextShape()
	shape.SetHeight(300).SetWidth(600).SetOffsetX(100).SetOffsetY(200)
	tr := shape.CreateTextRun("Hello World")
	tr.GetFont().SetBold(true).SetSize(24).SetColor(NewColor("FFE06B20"))

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, err := readFileFromZip(zr, "ppt/slides/slide1.xml")
	if err != nil {
		t.Fatalf("failed to read slide1.xml: %v", err)
	}

	content := string(data)
	if !strings.Contains(content, "Hello World") {
		t.Error("slide should contain text")
	}
	if !strings.Contains(content, `b="1"`) {
		t.Error("slide should contain bold attribute")
	}
	if !strings.Contains(content, `sz="2400"`) {
		t.Error("slide should contain font size")
	}
}

func TestWriteAutoShapeSlide(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	shape := slide.CreateAutoShape()
	shape.SetAutoShapeType(AutoShapeEllipse)
	shape.BaseShape.SetOffsetX(100).SetOffsetY(200).SetWidth(300).SetHeight(400)
	shape.GetFill().SetSolid(ColorBlue)
	shape.SetText("Inside text")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	if !strings.Contains(content, `prst="ellipse"`) {
		t.Error("slide should contain ellipse shape")
	}
	if !strings.Contains(content, "Inside text") {
		t.Error("slide should contain shape text")
	}
	if !strings.Contains(content, "solidFill") {
		t.Error("slide should contain solid fill")
	}
}

func TestWriteLineShapeSlide(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	line := slide.CreateLineShape()
	line.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(0)
	line.SetLineWidth(2).SetLineColor(ColorRed)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	if !strings.Contains(content, "cxnSp") {
		t.Error("slide should contain connection shape (line)")
	}
	if !strings.Contains(content, "FF0000") {
		t.Error("slide should contain red color")
	}
}

func TestWriteTableSlide(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	table := slide.CreateTableShape(2, 3)
	table.SetWidth(6000000).SetHeight(2000000)
	table.BaseShape.SetOffsetX(100000).SetOffsetY(100000)
	table.GetCell(0, 0).SetText("Header 1")
	table.GetCell(0, 1).SetText("Header 2")
	table.GetCell(0, 2).SetText("Header 3")
	table.GetCell(1, 0).SetText("Data 1")
	table.GetCell(1, 1).SetText("Data 2")
	table.GetCell(1, 2).SetText("Data 3")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	if !strings.Contains(content, "graphicFrame") {
		t.Error("slide should contain graphic frame for table")
	}
	if !strings.Contains(content, "Header 1") {
		t.Error("slide should contain table header text")
	}
	if !strings.Contains(content, "Data 3") {
		t.Error("slide should contain table data text")
	}
	if !strings.Contains(content, "gridCol") {
		t.Error("slide should contain grid columns")
	}
}

func TestWriteImageSlide(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// Create a minimal 1x1 PNG
	pngData := createMinimalPNG()

	img := slide.CreateDrawingShape()
	img.SetImageData(pngData, "image/png")
	img.SetHeight(100).SetWidth(100).SetOffsetX(50).SetOffsetY(50)
	img.BaseShape.SetName("Test Image")
	img.BaseShape.SetDescription("A test image")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Check image exists in media
	found := false
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/media/image") {
			found = true
			break
		}
	}
	if !found {
		t.Error("image file should exist in ppt/media/")
	}

	// Check content types include png
	data, _ := readFileFromZip(zr, "[Content_Types].xml")
	if !strings.Contains(string(data), "image/png") {
		t.Error("content types should include image/png")
	}

	// Check slide references image
	slideData, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	if !strings.Contains(string(slideData), "p:pic") {
		t.Error("slide should contain picture element")
	}
}

func TestWriteMultipleSlides(t *testing.T) {
	p := New()
	p.CreateSlide()
	p.CreateSlide()

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	for i := 1; i <= 3; i++ {
		path := "ppt/slides/slide" + strings.Replace("N", "N", string(rune('0'+i)), 1) + ".xml"
		_ = path
	}

	// Check all 3 slides exist
	slideCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			slideCount++
		}
	}
	if slideCount != 3 {
		t.Errorf("expected 3 slides, got %d", slideCount)
	}
}

func TestWriteToFile(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	shape := slide.CreateRichTextShape()
	shape.SetHeight(300).SetWidth(600).SetOffsetX(100).SetOffsetY(200)
	shape.CreateTextRun("File write test")

	tmpFile := "test_output.pptx"
	defer os.Remove(tmpFile)

	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.Save(tmpFile)
	if err != nil {
		t.Fatalf("Save error: %v", err)
	}

	// Verify file exists and is valid
	info, err := os.Stat(tmpFile)
	if err != nil {
		t.Fatalf("file not created: %v", err)
	}
	if info.Size() == 0 {
		t.Error("file is empty")
	}

	// Verify it's a valid zip
	f, _ := os.Open(tmpFile)
	defer f.Close()
	_, err = zip.NewReader(f, info.Size())
	if err != nil {
		t.Fatalf("output is not a valid zip: %v", err)
	}
}

func TestWriteXMLValidity(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// Add various shapes
	rt := slide.CreateRichTextShape()
	rt.SetHeight(300).SetWidth(600).SetOffsetX(100).SetOffsetY(200)
	rt.CreateTextRun("Test text with <special> & \"chars\"")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Verify all XML files are well-formed
	for _, f := range zr.File {
		if strings.HasSuffix(f.Name, ".xml") {
			rc, err := f.Open()
			if err != nil {
				t.Errorf("failed to open %s: %v", f.Name, err)
				continue
			}
			decoder := xml.NewDecoder(rc)
			for {
				_, err := decoder.Token()
				if err != nil {
					if err.Error() == "EOF" {
						break
					}
					t.Errorf("invalid XML in %s: %v", f.Name, err)
					break
				}
			}
			rc.Close()
		}
	}
}

func TestWriteRelationshipsValidity(t *testing.T) {
	p := New()
	p.CreateSlide()

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Check root rels
	data, _ := readFileFromZip(zr, "_rels/.rels")
	content := string(data)
	if !strings.Contains(content, "officeDocument") {
		t.Error("root rels should reference office document")
	}
	if !strings.Contains(content, "core-properties") {
		t.Error("root rels should reference core properties")
	}

	// Check presentation rels
	data, _ = readFileFromZip(zr, "ppt/_rels/presentation.xml.rels")
	content = string(data)
	if !strings.Contains(content, "slideMaster") {
		t.Error("presentation rels should reference slide master")
	}
	if !strings.Contains(content, "slide1.xml") {
		t.Error("presentation rels should reference slide1")
	}
	if !strings.Contains(content, "slide2.xml") {
		t.Error("presentation rels should reference slide2")
	}
	if !strings.Contains(content, "theme") {
		t.Error("presentation rels should reference theme")
	}
}

func TestWriteSlideMasterValidity(t *testing.T) {
	p := New()

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Check slide master
	data, _ := readFileFromZip(zr, "ppt/slideMasters/slideMaster1.xml")
	content := string(data)
	if !strings.Contains(content, "sldMaster") {
		t.Error("slide master should have sldMaster root element")
	}
	if !strings.Contains(content, "clrMap") {
		t.Error("slide master should have color map")
	}

	// Check slide master rels
	data, _ = readFileFromZip(zr, "ppt/slideMasters/_rels/slideMaster1.xml.rels")
	content = string(data)
	if !strings.Contains(content, "slideLayout") {
		t.Error("slide master rels should reference slide layout")
	}
	if !strings.Contains(content, "theme") {
		t.Error("slide master rels should reference theme")
	}
}

func TestWriteThemeValidity(t *testing.T) {
	p := New()

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	data, _ := readFileFromZip(zr, "ppt/theme/theme1.xml")
	content := string(data)
	if !strings.Contains(content, "theme") {
		t.Error("theme should have theme root element")
	}
	if !strings.Contains(content, "clrScheme") {
		t.Error("theme should have color scheme")
	}
	if !strings.Contains(content, "fontScheme") {
		t.Error("theme should have font scheme")
	}
	if !strings.Contains(content, "fmtScheme") {
		t.Error("theme should have format scheme")
	}
}

func TestWriteGradientFill(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	shape := slide.CreateAutoShape()
	shape.BaseShape.SetOffsetX(100).SetOffsetY(200).SetWidth(300).SetHeight(400)
	shape.GetFill().SetGradientLinear(ColorRed, ColorBlue, 90)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	if !strings.Contains(content, "gradFill") {
		t.Error("slide should contain gradient fill")
	}
}

func TestWriteViewProps(t *testing.T) {
	p := New()
	p.GetPresentationProperties().SetLastView(ViewNotes)
	p.GetPresentationProperties().SetZoom(2.0)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/viewProps.xml")
	content := string(data)

	if !strings.Contains(content, `lastView="notesView"`) {
		t.Error("view props should have notes view")
	}
	if !strings.Contains(content, `n="200"`) {
		t.Error("view props should have 200% zoom")
	}
}

// createMinimalPNG creates a minimal valid 1x1 white PNG image.
func createMinimalPNG() []byte {
	return []byte{
		0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
		0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
		0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
		0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
		0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, // IDAT chunk
		0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
		0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
		0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, // IEND chunk
		0x44, 0xAE, 0x42, 0x60, 0x82,
	}
}
