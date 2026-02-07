package gopresentation

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"os"
	"strings"
	"testing"
	"time"
)

// ===== Comment Tests =====

func TestCommentTypes(t *testing.T) {
	author := NewCommentAuthor("John Doe", "JD")
	if author.Name != "John Doe" {
		t.Error("author name mismatch")
	}
	if author.Initials != "JD" {
		t.Error("author initials mismatch")
	}

	c := NewComment()
	c.SetAuthor(author).SetText("Test comment").SetPosition(100, 200)
	if c.Text != "Test comment" {
		t.Error("comment text mismatch")
	}
	if c.PositionX != 100 || c.PositionY != 200 {
		t.Error("comment position mismatch")
	}
	if c.Author != author {
		t.Error("comment author mismatch")
	}

	now := time.Now()
	c.SetDate(now)
	if c.Date != now {
		t.Error("comment date mismatch")
	}
}

func TestSlideComments(t *testing.T) {
	s := newSlide()
	if s.GetCommentCount() != 0 {
		t.Error("expected 0 comments")
	}

	author := NewCommentAuthor("Alice", "A")
	c1 := NewComment().SetAuthor(author).SetText("First")
	c2 := NewComment().SetAuthor(author).SetText("Second")
	s.AddComment(c1)
	s.AddComment(c2)

	if s.GetCommentCount() != 2 {
		t.Errorf("expected 2 comments, got %d", s.GetCommentCount())
	}
	comments := s.GetComments()
	if comments[0].Text != "First" || comments[1].Text != "Second" {
		t.Error("comment text mismatch")
	}
}

func TestWriteCommentsRoundTrip(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	author := NewCommentAuthor("Bob", "B")
	c := NewComment().SetAuthor(author).SetText("Review this slide").SetPosition(50, 75)
	c.SetDate(time.Date(2025, 6, 15, 10, 30, 0, 0, time.UTC))
	slide.AddComment(c)

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Verify comment authors file exists
	authData, err := readFileFromZip(zr, "ppt/commentAuthors.xml")
	if err != nil {
		t.Fatal("commentAuthors.xml not found")
	}
	if !strings.Contains(string(authData), "Bob") {
		t.Error("comment authors should contain Bob")
	}

	// Verify comment file exists
	cmData, err := readFileFromZip(zr, "ppt/comments/comment1.xml")
	if err != nil {
		t.Fatal("comment1.xml not found")
	}
	cmContent := string(cmData)
	if !strings.Contains(cmContent, "Review this slide") {
		t.Error("comment should contain text")
	}
	if !strings.Contains(cmContent, `x="50"`) {
		t.Error("comment should contain position x")
	}

	// Verify content types include comments
	ctData, _ := readFileFromZip(zr, "[Content_Types].xml")
	ctContent := string(ctData)
	if !strings.Contains(ctContent, "commentAuthors") {
		t.Error("content types should include commentAuthors")
	}
	if !strings.Contains(ctContent, "comments") {
		t.Error("content types should include comments")
	}

	// Verify slide rels reference comments
	relData, _ := readFileFromZip(zr, "ppt/slides/_rels/slide1.xml.rels")
	if !strings.Contains(string(relData), "comments") {
		t.Error("slide rels should reference comments")
	}

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	readSlide := pres.GetActiveSlide()
	if readSlide.GetCommentCount() != 1 {
		t.Fatalf("expected 1 comment, got %d", readSlide.GetCommentCount())
	}
	readComment := readSlide.GetComments()[0]
	if readComment.Text != "Review this slide" {
		t.Errorf("expected comment text 'Review this slide', got '%s'", readComment.Text)
	}
	if readComment.PositionX != 50 || readComment.PositionY != 75 {
		t.Errorf("expected position (50,75), got (%d,%d)", readComment.PositionX, readComment.PositionY)
	}
}

// ===== Bullet Tests =====

func TestBulletTypes(t *testing.T) {
	b := NewBullet()
	if b.Type != BulletTypeNone {
		t.Error("default bullet type should be none")
	}
	if b.StartAt != 1 {
		t.Error("default start at should be 1")
	}
	if b.Size != 100 {
		t.Error("default size should be 100")
	}

	b.SetCharBullet("•", "Arial")
	if b.Type != BulletTypeChar {
		t.Error("expected char bullet type")
	}
	if b.Style != "•" {
		t.Error("expected bullet char •")
	}
	if b.Font != "Arial" {
		t.Error("expected font Arial")
	}

	b2 := NewBullet()
	b2.SetNumericBullet(NumFormatArabicPeriod, 5)
	if b2.Type != BulletTypeNumeric {
		t.Error("expected numeric bullet type")
	}
	if b2.NumFormat != NumFormatArabicPeriod {
		t.Error("expected arabicPeriod format")
	}
	if b2.StartAt != 5 {
		t.Error("expected start at 5")
	}

	b3 := NewBullet()
	b3.SetColor(ColorRed).SetSize(150)
	if b3.Color == nil || b3.Color.ARGB != "FFFF0000" {
		t.Error("expected red color")
	}
	if b3.Size != 150 {
		t.Error("expected size 150")
	}

	// Test size clamping
	b4 := NewBullet()
	b4.SetSize(10) // below min
	if b4.Size != 25 {
		t.Errorf("expected size clamped to 25, got %d", b4.Size)
	}
	b4.SetSize(500) // above max
	if b4.Size != 400 {
		t.Errorf("expected size clamped to 400, got %d", b4.Size)
	}
}

func TestWriteBulletRoundTrip(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetHeight(500).SetWidth(800).SetOffsetX(100).SetOffsetY(100)

	// Paragraph with character bullet
	para1 := rt.GetActiveParagraph()
	bullet1 := NewBullet().SetCharBullet("•", "Arial")
	bullet1.SetColor(ColorRed).SetSize(120)
	para1.SetBullet(bullet1)
	para1.CreateTextRun("First bullet item")

	// Paragraph with numeric bullet
	para2 := rt.CreateParagraph()
	bullet2 := NewBullet().SetNumericBullet(NumFormatArabicPeriod, 1)
	para2.SetBullet(bullet2)
	para2.CreateTextRun("Numbered item")

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	slideData, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(slideData)

	if !strings.Contains(content, "buChar") {
		t.Error("slide should contain buChar element")
	}
	if !strings.Contains(content, "buAutoNum") {
		t.Error("slide should contain buAutoNum element")
	}
	if !strings.Contains(content, "buFont") {
		t.Error("slide should contain buFont element")
	}
	if !strings.Contains(content, "buClr") {
		t.Error("slide should contain buClr element")
	}
	if !strings.Contains(content, "buSzPct") {
		t.Error("slide should contain buSzPct element")
	}

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	shapes := pres.GetActiveSlide().GetShapes()
	var rtShape *RichTextShape
	for _, s := range shapes {
		if r, ok := s.(*RichTextShape); ok {
			rtShape = r
			break
		}
	}
	if rtShape == nil {
		t.Fatal("expected RichTextShape")
	}

	paras := rtShape.GetParagraphs()
	if len(paras) < 2 {
		t.Fatalf("expected at least 2 paragraphs, got %d", len(paras))
	}

	// Check first paragraph bullet
	b1 := paras[0].GetBullet()
	if b1 == nil {
		t.Fatal("expected bullet on first paragraph")
	}
	if b1.Type != BulletTypeChar {
		t.Errorf("expected char bullet, got %d", b1.Type)
	}
	if b1.Style != "•" {
		t.Errorf("expected bullet char •, got %s", b1.Style)
	}
	if b1.Font != "Arial" {
		t.Errorf("expected font Arial, got %s", b1.Font)
	}
	if b1.Color == nil || b1.Color.ARGB != "FFFF0000" {
		t.Error("expected red bullet color")
	}
	if b1.Size != 120 {
		t.Errorf("expected bullet size 120, got %d", b1.Size)
	}

	// Check second paragraph bullet
	b2 := paras[1].GetBullet()
	if b2 == nil {
		t.Fatal("expected bullet on second paragraph")
	}
	if b2.Type != BulletTypeNumeric {
		t.Errorf("expected numeric bullet, got %d", b2.Type)
	}
	if b2.NumFormat != NumFormatArabicPeriod {
		t.Errorf("expected arabicPeriod, got %s", b2.NumFormat)
	}
}

// ===== Group Shape Tests =====

func TestGroupShapeTypes(t *testing.T) {
	g := NewGroupShape()
	if g.GetType() != ShapeTypeGroup {
		t.Error("expected ShapeTypeGroup")
	}
	if g.GetShapeCount() != 0 {
		t.Error("expected 0 shapes")
	}

	rt := NewRichTextShape()
	g.AddShape(rt)
	if g.GetShapeCount() != 1 {
		t.Error("expected 1 shape")
	}

	as := NewAutoShape()
	g.AddShape(as)
	if g.GetShapeCount() != 2 {
		t.Error("expected 2 shapes")
	}

	err := g.RemoveShape(0)
	if err != nil {
		t.Errorf("unexpected error: %v", err)
	}
	if g.GetShapeCount() != 1 {
		t.Error("expected 1 shape after remove")
	}

	err = g.RemoveShape(5)
	if err == nil {
		t.Error("expected error for out of range")
	}
}

func TestWriteGroupShapeRoundTrip(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	group := slide.CreateGroupShape()
	group.BaseShape.SetOffsetX(100).SetOffsetY(200).SetWidth(5000000).SetHeight(3000000)
	group.BaseShape.SetName("Test Group")

	rt := NewRichTextShape()
	rt.SetOffsetX(100).SetOffsetY(200).SetWidth(2000000).SetHeight(500000)
	rt.CreateTextRun("Inside group")
	group.AddShape(rt)

	line := NewLineShape()
	line.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(1000000).SetHeight(0)
	group.AddShape(line)

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	slideData, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(slideData)

	if !strings.Contains(content, "grpSp") {
		t.Error("slide should contain grpSp element")
	}
	if !strings.Contains(content, "Inside group") {
		t.Error("slide should contain group child text")
	}
	if !strings.Contains(content, "chOff") {
		t.Error("slide should contain child offset")
	}

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	shapes := pres.GetActiveSlide().GetShapes()
	var grp *GroupShape
	for _, s := range shapes {
		if g, ok := s.(*GroupShape); ok {
			grp = g
			break
		}
	}
	if grp == nil {
		t.Fatal("expected GroupShape")
	}
	if grp.GetShapeCount() < 1 {
		t.Errorf("expected at least 1 child shape, got %d", grp.GetShapeCount())
	}
}

// ===== Placeholder Shape Tests =====

func TestPlaceholderShapeTypes(t *testing.T) {
	ph := NewPlaceholderShape(PlaceholderTitle)
	if ph.GetType() != ShapeTypePlaceholder {
		t.Error("expected ShapeTypePlaceholder")
	}
	if ph.GetPlaceholderType() != PlaceholderTitle {
		t.Error("expected title placeholder type")
	}

	ph.SetPlaceholderIndex(5)
	if ph.GetPlaceholderIndex() != 5 {
		t.Error("expected placeholder index 5")
	}

	// Test all placeholder types
	types := []PlaceholderType{
		PlaceholderTitle, PlaceholderBody, PlaceholderCtrTitle,
		PlaceholderSubTitle, PlaceholderDate, PlaceholderFooter, PlaceholderSlideNum,
	}
	for _, pt := range types {
		p := NewPlaceholderShape(pt)
		if p.GetPlaceholderType() != pt {
			t.Errorf("placeholder type mismatch for %s", pt)
		}
	}
}

func TestWritePlaceholderRoundTrip(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	ph := slide.CreatePlaceholderShape(PlaceholderTitle)
	ph.BaseShape.SetOffsetX(500000).SetOffsetY(300000).SetWidth(8000000).SetHeight(1000000)
	ph.CreateTextRun("Slide Title")

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	slideData, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(slideData)

	if !strings.Contains(content, `ph type="title"`) {
		t.Error("slide should contain placeholder type")
	}
	if !strings.Contains(content, "Slide Title") {
		t.Error("slide should contain placeholder text")
	}
	if !strings.Contains(content, "noGrp") {
		t.Error("placeholder should have noGrp lock")
	}

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	shapes := pres.GetActiveSlide().GetShapes()
	var readPh *PlaceholderShape
	for _, s := range shapes {
		if p, ok := s.(*PlaceholderShape); ok {
			readPh = p
			break
		}
	}
	if readPh == nil {
		t.Fatal("expected PlaceholderShape")
	}
	if readPh.GetPlaceholderType() != PlaceholderTitle {
		t.Errorf("expected title placeholder, got %s", readPh.GetPlaceholderType())
	}

	paras := readPh.GetParagraphs()
	if len(paras) == 0 {
		t.Fatal("expected at least 1 paragraph")
	}
	elems := paras[0].GetElements()
	if len(elems) == 0 {
		t.Fatal("expected at least 1 element")
	}
	if tr, ok := elems[0].(*TextRun); ok {
		if tr.GetText() != "Slide Title" {
			t.Errorf("expected 'Slide Title', got '%s'", tr.GetText())
		}
	}
}

// ===== Notes Slide Tests =====

func TestWriteNotesRoundTrip(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	slide.SetNotes("These are speaker notes for the first slide.")

	rt := slide.CreateRichTextShape()
	rt.SetHeight(100).SetWidth(200)
	rt.CreateTextRun("Content")

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Verify notes slide exists
	notesData, err := readFileFromZip(zr, "ppt/notesSlides/notesSlide1.xml")
	if err != nil {
		t.Fatal("notesSlide1.xml not found")
	}
	notesContent := string(notesData)
	if !strings.Contains(notesContent, "speaker notes") {
		t.Error("notes should contain text")
	}
	if !strings.Contains(notesContent, "p:notes") {
		t.Error("notes should have p:notes root")
	}

	// Verify notes rels
	notesRels, err := readFileFromZip(zr, "ppt/notesSlides/_rels/notesSlide1.xml.rels")
	if err != nil {
		t.Fatal("notes rels not found")
	}
	if !strings.Contains(string(notesRels), "slide1.xml") {
		t.Error("notes rels should reference parent slide")
	}

	// Verify content types include notes
	ctData, _ := readFileFromZip(zr, "[Content_Types].xml")
	if !strings.Contains(string(ctData), "notesSlide") {
		t.Error("content types should include notesSlide")
	}

	// Verify slide rels reference notes
	slideRels, _ := readFileFromZip(zr, "ppt/slides/_rels/slide1.xml.rels")
	if !strings.Contains(string(slideRels), "notesSlide") {
		t.Error("slide rels should reference notes slide")
	}

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	readSlide := pres.GetActiveSlide()
	if readSlide.GetNotes() != "These are speaker notes for the first slide." {
		t.Errorf("expected notes text, got '%s'", readSlide.GetNotes())
	}
}

func TestWriteMultiSlideNotes(t *testing.T) {
	p := New()
	slide1 := p.GetActiveSlide()
	slide1.SetNotes("Notes for slide 1")

	slide2 := p.CreateSlide()
	// slide2 has no notes

	slide3 := p.CreateSlide()
	slide3.SetNotes("Notes for slide 3")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Slide 1 should have notes
	_, err = readFileFromZip(zr, "ppt/notesSlides/notesSlide1.xml")
	if err != nil {
		t.Error("notesSlide1.xml should exist")
	}

	// Slide 3 should have notes
	_, err = readFileFromZip(zr, "ppt/notesSlides/notesSlide3.xml")
	if err != nil {
		t.Error("notesSlide3.xml should exist")
	}

	_ = slide2
}

// ===== Background Tests =====

func TestWriteBackgroundRoundTrip(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	bg := NewFill().SetSolid(ColorBlue)
	slide.SetBackground(bg)

	rt := slide.CreateRichTextShape()
	rt.SetHeight(100).SetWidth(200)
	rt.CreateTextRun("With background")

	// Write
	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	slideData, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(slideData)

	if !strings.Contains(content, "p:bg") {
		t.Error("slide should contain p:bg element")
	}
	if !strings.Contains(content, "bgPr") {
		t.Error("slide should contain bgPr element")
	}
	if !strings.Contains(content, "0000FF") {
		t.Error("slide should contain blue color")
	}

	// Read back
	reader := &PPTXReader{}
	data := buf.Bytes()
	pres, err := reader.ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	readSlide := pres.GetActiveSlide()
	readBg := readSlide.GetBackground()
	if readBg == nil {
		t.Fatal("expected background fill")
	}
	if readBg.Type != FillSolid {
		t.Errorf("expected solid fill, got %d", readBg.Type)
	}
	if readBg.Color.ARGB != "FF0000FF" {
		t.Errorf("expected blue color, got %s", readBg.Color.ARGB)
	}
}

// ===== Animation Tests =====

func TestAnimationTypes(t *testing.T) {
	a := NewAnimation()
	if len(a.GetShapeIndexes()) != 0 {
		t.Error("expected 0 shape indexes")
	}

	a.AddShape(0).AddShape(1).AddShape(2)
	if len(a.GetShapeIndexes()) != 3 {
		t.Errorf("expected 3 shape indexes, got %d", len(a.GetShapeIndexes()))
	}
}

func TestSlideAnimations(t *testing.T) {
	s := newSlide()
	a := NewAnimation()
	a.AddShape(0)
	s.AddAnimation(a)

	if len(s.GetAnimations()) != 1 {
		t.Error("expected 1 animation")
	}
}

// ===== Paragraph Spacing Tests =====

func TestParagraphSpacing(t *testing.T) {
	p := NewParagraph()
	p.SetSpaceBefore(500)
	p.SetSpaceAfter(300)
	p.SetLineSpacing(200)

	if p.GetSpaceBefore() != 500 {
		t.Errorf("expected space before 500, got %d", p.GetSpaceBefore())
	}
	if p.GetSpaceAfter() != 300 {
		t.Errorf("expected space after 300, got %d", p.GetSpaceAfter())
	}
	if p.GetLineSpacing() != 200 {
		t.Errorf("expected line spacing 200, got %d", p.GetLineSpacing())
	}
}

// ===== OOXML Compliance Tests for New Features =====

func TestOOXMLChartCompliance(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(5000000).SetHeight(3000000)
	chart.GetTitle().SetText("Compliance Chart").SetVisible(true)
	chart.GetLegend().Visible = true
	chart.GetLegend().Position = LegendBottom

	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("Data", []string{"A", "B", "C"}, []float64{10, 20, 30}))
	chart.GetPlotArea().SetType(bar)
	chart.GetPlotArea().GetAxisX().SetTitle("X")
	chart.GetPlotArea().GetAxisY().SetTitle("Y")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Chart XML must be well-formed
	chartData, err := readFileFromZip(zr, "ppt/charts/chart1.xml")
	if err != nil {
		t.Fatal("chart1.xml not found")
	}

	decoder := xml.NewDecoder(strings.NewReader(string(chartData)))
	for {
		_, err := decoder.Token()
		if err != nil {
			if err.Error() == "EOF" {
				break
			}
			t.Fatalf("chart XML is malformed: %v", err)
		}
	}

	content := string(chartData)

	// Must have chartSpace root
	if !strings.Contains(content, "c:chartSpace") {
		t.Error("chart must have c:chartSpace root")
	}

	// Must have correct namespaces
	if !strings.Contains(content, "http://schemas.openxmlformats.org/drawingml/2006/chart") {
		t.Error("chart must have chart namespace")
	}

	// Must have chart element
	if !strings.Contains(content, "c:chart") {
		t.Error("chart must have c:chart element")
	}

	// Must have plotArea
	if !strings.Contains(content, "c:plotArea") {
		t.Error("chart must have c:plotArea")
	}

	// Must have layout
	if !strings.Contains(content, "c:layout") {
		t.Error("chart must have c:layout")
	}

	// Must have title
	if !strings.Contains(content, "c:title") {
		t.Error("chart must have c:title")
	}

	// Must have legend
	if !strings.Contains(content, "c:legend") {
		t.Error("chart must have c:legend")
	}

	// Must have axes for bar chart
	if !strings.Contains(content, "c:catAx") {
		t.Error("bar chart must have category axis")
	}
	if !strings.Contains(content, "c:valAx") {
		t.Error("bar chart must have value axis")
	}

	// Must have axis IDs
	if !strings.Contains(content, "c:axId") {
		t.Error("chart must have axis IDs")
	}

	// Must have series
	if !strings.Contains(content, "c:ser") {
		t.Error("chart must have series")
	}

	// Must have plotVisOnly
	if !strings.Contains(content, "c:plotVisOnly") {
		t.Error("chart must have plotVisOnly")
	}

	// Must have dispBlanksAs
	if !strings.Contains(content, "c:dispBlanksAs") {
		t.Error("chart must have dispBlanksAs")
	}
}

func TestOOXMLCommentCompliance(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	author := NewCommentAuthor("Test Author", "TA")
	c := NewComment().SetAuthor(author).SetText("Compliance comment").SetPosition(10, 20)
	c.SetDate(time.Date(2025, 1, 1, 0, 0, 0, 0, time.UTC))
	slide.AddComment(c)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Comment authors must be well-formed XML
	authData, _ := readFileFromZip(zr, "ppt/commentAuthors.xml")
	decoder := xml.NewDecoder(strings.NewReader(string(authData)))
	for {
		_, err := decoder.Token()
		if err != nil {
			if err.Error() == "EOF" {
				break
			}
			t.Fatalf("commentAuthors XML is malformed: %v", err)
		}
	}

	authContent := string(authData)
	if !strings.Contains(authContent, "p:cmAuthorLst") {
		t.Error("comment authors must have p:cmAuthorLst root")
	}
	if !strings.Contains(authContent, nsPresentationML) {
		t.Error("comment authors must have PresentationML namespace")
	}
	if !strings.Contains(authContent, "p:cmAuthor") {
		t.Error("comment authors must have p:cmAuthor elements")
	}

	// Comments must be well-formed XML
	cmData, _ := readFileFromZip(zr, "ppt/comments/comment1.xml")
	decoder2 := xml.NewDecoder(strings.NewReader(string(cmData)))
	for {
		_, err := decoder2.Token()
		if err != nil {
			if err.Error() == "EOF" {
				break
			}
			t.Fatalf("comment XML is malformed: %v", err)
		}
	}

	cmContent := string(cmData)
	if !strings.Contains(cmContent, "p:cmLst") {
		t.Error("comments must have p:cmLst root")
	}
	if !strings.Contains(cmContent, "p:cm") {
		t.Error("comments must have p:cm elements")
	}
	if !strings.Contains(cmContent, "p:pos") {
		t.Error("comments must have p:pos element")
	}
	if !strings.Contains(cmContent, "p:text") {
		t.Error("comments must have p:text element")
	}

	// Presentation rels must reference comment authors
	presRels, _ := readFileFromZip(zr, "ppt/_rels/presentation.xml.rels")
	if !strings.Contains(string(presRels), "commentAuthors") {
		t.Error("presentation rels must reference commentAuthors")
	}
}

func TestOOXMLNotesSlideCompliance(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	slide.SetNotes("Compliance notes text")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Notes slide must be well-formed XML
	notesData, _ := readFileFromZip(zr, "ppt/notesSlides/notesSlide1.xml")
	decoder := xml.NewDecoder(strings.NewReader(string(notesData)))
	for {
		_, err := decoder.Token()
		if err != nil {
			if err.Error() == "EOF" {
				break
			}
			t.Fatalf("notes slide XML is malformed: %v", err)
		}
	}

	content := string(notesData)

	// Must have p:notes root
	if !strings.Contains(content, "p:notes") {
		t.Error("notes slide must have p:notes root")
	}

	// Must have correct namespaces
	if !strings.Contains(content, nsPresentationML) {
		t.Error("notes slide must have PresentationML namespace")
	}
	if !strings.Contains(content, nsDrawingML) {
		t.Error("notes slide must have DrawingML namespace")
	}

	// Must have cSld
	if !strings.Contains(content, "p:cSld") {
		t.Error("notes slide must have p:cSld")
	}

	// Must have spTree
	if !strings.Contains(content, "p:spTree") {
		t.Error("notes slide must have p:spTree")
	}

	// Must have placeholder with body type
	if !strings.Contains(content, `ph type="body"`) {
		t.Error("notes slide must have body placeholder")
	}

	// Notes rels must reference parent slide
	notesRels, _ := readFileFromZip(zr, "ppt/notesSlides/_rels/notesSlide1.xml.rels")
	if !strings.Contains(string(notesRels), relTypeSlide) {
		t.Error("notes rels must reference parent slide")
	}
}

// ===== Combined Feature Test =====

func TestWriteAllFeaturesXMLWellFormed(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	slide.SetNotes("Speaker notes")
	slide.SetBackground(NewFill().SetSolid(ColorWhite))

	// Rich text with bullet
	rt := slide.CreateRichTextShape()
	rt.SetHeight(500).SetWidth(800).SetOffsetX(100).SetOffsetY(100)
	para := rt.GetActiveParagraph()
	para.SetBullet(NewBullet().SetCharBullet("•"))
	para.CreateTextRun("Bullet item")

	// Placeholder
	ph := slide.CreatePlaceholderShape(PlaceholderTitle)
	ph.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(9144000).SetHeight(1000000)
	ph.CreateTextRun("Title")

	// Chart
	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(100).SetOffsetY(2000000).SetWidth(5000000).SetHeight(3000000)
	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("S1", []string{"A"}, []float64{10}))
	chart.GetPlotArea().SetType(bar)

	// Group
	group := slide.CreateGroupShape()
	group.BaseShape.SetOffsetX(6000000).SetOffsetY(2000000).SetWidth(3000000).SetHeight(2000000)
	childRT := NewRichTextShape()
	childRT.SetOffsetX(0).SetOffsetY(0).SetWidth(1000000).SetHeight(500000)
	childRT.CreateTextRun("In group")
	group.AddShape(childRT)

	// Comment
	author := NewCommentAuthor("Tester", "T")
	slide.AddComment(NewComment().SetAuthor(author).SetText("Check this"))

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Verify ALL XML files are well-formed
	xmlCount := 0
	for _, f := range zr.File {
		if strings.HasSuffix(f.Name, ".xml") || strings.HasSuffix(f.Name, ".rels") {
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
					t.Errorf("malformed XML in %s: %v", f.Name, err)
					break
				}
			}
			rc.Close()
			xmlCount++
		}
	}

	if xmlCount < 15 {
		t.Errorf("expected at least 15 XML files, got %d", xmlCount)
	}
	t.Logf("validated %d XML files for well-formedness (all features)", xmlCount)
}

// ===== Coverage Expansion Tests =====

// --- 0% coverage: chart.go GetView3D(), PlotArea.GetType() ---

func TestChartGetView3D(t *testing.T) {
	chart := NewChartShape()
	v := chart.GetView3D()
	if v == nil {
		t.Fatal("GetView3D should not return nil")
	}
	if v.RotX != 15 {
		t.Errorf("expected RotX 15, got %d", v.RotX)
	}
	if v.RotY != 20 {
		t.Errorf("expected RotY 20, got %d", v.RotY)
	}
	if v.DepthPercent != 100 {
		t.Errorf("expected DepthPercent 100, got %d", v.DepthPercent)
	}
	if v.HeightPercent == nil || *v.HeightPercent != 100 {
		t.Error("expected HeightPercent 100")
	}
	if !v.RightAngleAxes {
		t.Error("expected RightAngleAxes true")
	}

	// Test SetHeightPercent with nil (autoscale)
	v.SetHeightPercent(nil)
	if v.HeightPercent != nil {
		t.Error("expected HeightPercent nil after SetHeightPercent(nil)")
	}
}

func TestPlotAreaGetType(t *testing.T) {
	pa := NewPlotArea()
	if pa.GetType() != nil {
		t.Error("expected nil chart type for new PlotArea")
	}

	bar := NewBarChart()
	pa.SetType(bar)
	ct := pa.GetType()
	if ct == nil {
		t.Fatal("expected non-nil chart type")
	}
	if ct.GetChartTypeName() != "bar" {
		t.Errorf("expected 'bar', got '%s'", ct.GetChartTypeName())
	}
}

// --- 0% coverage: group.go GetShapes(), errorString.Error() ---

func TestGroupShapeGetShapes(t *testing.T) {
	g := NewGroupShape()
	shapes := g.GetShapes()
	if len(shapes) != 0 {
		t.Error("expected empty shapes slice")
	}

	rt := NewRichTextShape()
	as := NewAutoShape()
	g.AddShape(rt)
	g.AddShape(as)

	shapes = g.GetShapes()
	if len(shapes) != 2 {
		t.Errorf("expected 2 shapes, got %d", len(shapes))
	}
	if shapes[0].GetType() != ShapeTypeRichText {
		t.Error("first shape should be RichText")
	}
	if shapes[1].GetType() != ShapeTypeAutoShape {
		t.Error("second shape should be AutoShape")
	}
}

func TestErrorStringError(t *testing.T) {
	// Test the errorString type via RemoveShape out of range
	g := NewGroupShape()
	err := g.RemoveShape(0)
	if err == nil {
		t.Fatal("expected error")
	}
	if err.Error() != "index out of range" {
		t.Errorf("expected 'index out of range', got '%s'", err.Error())
	}

	// Negative index
	err = g.RemoveShape(-1)
	if err == nil {
		t.Fatal("expected error for negative index")
	}
}

// --- 0% coverage: shape.go SetFill, SetBorder, SetShadow, SetAlignment ---

func TestBaseShapeSetFillBorderShadow(t *testing.T) {
	bs := &BaseShape{}

	// SetFill
	fill := NewFill().SetSolid(ColorRed)
	bs.SetFill(fill)
	if bs.GetFill().Type != FillSolid {
		t.Error("expected solid fill")
	}
	if bs.GetFill().Color.ARGB != "FFFF0000" {
		t.Error("expected red fill color")
	}

	// SetBorder
	border := &Border{Style: BorderSolid, Width: 2, Color: ColorBlue}
	bs.SetBorder(border)
	if bs.GetBorder().Style != BorderSolid {
		t.Error("expected solid border")
	}
	if bs.GetBorder().Width != 2 {
		t.Error("expected border width 2")
	}

	// SetShadow
	shadow := NewShadow()
	shadow.SetVisible(true).SetDirection(45).SetDistance(5)
	bs.SetShadow(shadow)
	if !bs.GetShadow().Visible {
		t.Error("expected visible shadow")
	}
	if bs.GetShadow().Direction != 45 {
		t.Error("expected shadow direction 45")
	}
}

func TestParagraphSetAlignment(t *testing.T) {
	p := NewParagraph()
	align := NewAlignment().SetHorizontal(HorizontalCenter).SetVertical(VerticalMiddle)
	p.SetAlignment(align)

	got := p.GetAlignment()
	if got.Horizontal != HorizontalCenter {
		t.Errorf("expected center, got %s", got.Horizontal)
	}
	if got.Vertical != VerticalMiddle {
		t.Errorf("expected middle, got %s", got.Vertical)
	}
}

// --- 0% coverage: writer.go nextRelID() ---

func TestNextRelID(t *testing.T) {
	w := &PPTXWriter{presentation: New()}
	w.relID = 0
	id1 := w.nextRelID()
	if id1 != "rId1" {
		t.Errorf("expected rId1, got %s", id1)
	}
	id2 := w.nextRelID()
	if id2 != "rId2" {
		t.Errorf("expected rId2, got %s", id2)
	}
	id3 := w.nextRelID()
	if id3 != "rId3" {
		t.Errorf("expected rId3, got %s", id3)
	}
}

// --- writer_chart.go: Bar3D and Pie3D at 0% writer coverage ---

func TestWriteBar3DChart(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(5000000).SetHeight(3000000)
	chart.GetTitle().SetText("3D Bar Chart").SetVisible(true)

	bar3d := NewBar3DChart()
	bar3d.AddSeries(NewChartSeriesOrdered("Sales", []string{"Q1", "Q2", "Q3"}, []float64{100, 200, 150}))
	chart.GetPlotArea().SetType(bar3d)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	chartData, err := readFileFromZip(zr, "ppt/charts/chart1.xml")
	if err != nil {
		t.Fatal("chart1.xml not found")
	}
	content := string(chartData)
	if !strings.Contains(content, "c:bar3DChart") {
		t.Error("expected c:bar3DChart element")
	}
	if !strings.Contains(content, "c:barDir") {
		t.Error("expected c:barDir element")
	}
	if !strings.Contains(content, "c:gapWidth") {
		t.Error("expected c:gapWidth element")
	}
}

func TestWritePie3DChart(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(5000000).SetHeight(3000000)
	chart.GetTitle().SetText("3D Pie Chart").SetVisible(true)

	pie3d := NewPie3DChart()
	s := NewChartSeriesOrdered("Market", []string{"A", "B", "C"}, []float64{40, 35, 25})
	s.ShowPercentage = true
	s.ShowValue = true
	pie3d.AddSeries(s)
	chart.GetPlotArea().SetType(pie3d)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	chartData, err := readFileFromZip(zr, "ppt/charts/chart1.xml")
	if err != nil {
		t.Fatal("chart1.xml not found")
	}
	content := string(chartData)
	if !strings.Contains(content, "c:pie3DChart") {
		t.Error("expected c:pie3DChart element")
	}
	if !strings.Contains(content, "c:varyColors") {
		t.Error("expected c:varyColors element")
	}
	// Pie3D should NOT have axes
	if strings.Contains(content, "c:catAx") {
		t.Error("pie3D chart should not have category axis")
	}
}

// --- writer_xml.go: getImageExtension/getImageContentType coverage ---

func TestGetImageExtensionAndContentType(t *testing.T) {
	w := &PPTXWriter{presentation: New()}

	tests := []struct {
		mimeType    string
		path        string
		expectedExt string
		expectedCT  string
	}{
		{"image/png", "", "png", "image/png"},
		{"image/jpeg", "", "jpeg", "image/jpeg"},
		{"image/gif", "", "gif", "image/gif"},
		{"image/bmp", "", "bmp", "image/bmp"},
		{"image/svg+xml", "", "svg", "image/svg+xml"},
		{"", "photo.jpg", "jpeg", "image/jpeg"},
		{"", "photo.png", "png", "image/png"},
		{"", "photo.gif", "gif", "image/gif"},
		{"", "photo.bmp", "bmp", "image/bmp"},
		{"", "photo.svg", "svg", "image/svg+xml"},
		{"", "", "png", "image/png"}, // default
	}

	for _, tt := range tests {
		ds := &DrawingShape{}
		ds.mimeType = tt.mimeType
		ds.path = tt.path

		ext := w.getImageExtension(ds)
		if ext != tt.expectedExt {
			t.Errorf("mime=%q path=%q: expected ext %q, got %q", tt.mimeType, tt.path, tt.expectedExt, ext)
		}

		ct := w.getImageContentType(ds)
		if ct != tt.expectedCT {
			t.Errorf("mime=%q path=%q: expected ct %q, got %q", tt.mimeType, tt.path, tt.expectedCT, ct)
		}
	}
}

// --- writer_parts.go: writeViewProps with all view types ---

func TestWriteViewPropsAllViews(t *testing.T) {
	views := []struct {
		view     ViewType
		expected string
	}{
		{ViewSlide, "sldView"},
		{ViewNotes, "notesView"},
		{ViewHandout, "handoutView"},
		{ViewOutline, "outlineView"},
		{ViewSlideMaster, "sldMasterView"},
		{ViewSlideSorter, "sldSorterView"},
	}

	for _, v := range views {
		p := New()
		p.GetPresentationProperties().SetLastView(v.view)

		var buf bytes.Buffer
		w, _ := NewWriter(p, WriterPowerPoint2007)
		err := w.WriteTo(&buf)
		if err != nil {
			t.Fatalf("WriteTo error for view %d: %v", v.view, err)
		}

		zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
		data, err := readFileFromZip(zr, "ppt/viewProps.xml")
		if err != nil {
			t.Fatalf("viewProps.xml not found for view %d", v.view)
		}
		content := string(data)
		if !strings.Contains(content, fmt.Sprintf(`lastView="%s"`, v.expected)) {
			t.Errorf("expected lastView=%q in viewProps for view %d, got:\n%s", v.expected, v.view, content)
		}
	}
}

// --- writer_parts.go: writePresProps with Browse and Kiosk ---

func TestWritePresPropsAllSlideshowTypes(t *testing.T) {
	types := []struct {
		ssType   SlideshowType
		expected string
	}{
		{SlideshowTypePresent, "p:present"},
		{SlideshowTypeBrowse, "p:browse"},
		{SlideshowTypeKiosk, "p:kiosk"},
	}

	for _, tt := range types {
		p := New()
		p.GetPresentationProperties().SetSlideshowType(tt.ssType)

		var buf bytes.Buffer
		w, _ := NewWriter(p, WriterPowerPoint2007)
		err := w.WriteTo(&buf)
		if err != nil {
			t.Fatalf("WriteTo error for type %d: %v", tt.ssType, err)
		}

		zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
		data, err := readFileFromZip(zr, "ppt/presProps.xml")
		if err != nil {
			t.Fatalf("presProps.xml not found for type %d", tt.ssType)
		}
		content := string(data)
		if !strings.Contains(content, tt.expected) {
			t.Errorf("expected %q in presProps for type %d", tt.expected, tt.ssType)
		}
	}
}

// --- properties.go: SetLayout all types ---

func TestSetLayoutAllTypes(t *testing.T) {
	layouts := []struct {
		name     string
		expectedCX int64
	}{
		{LayoutScreen4x3, 9144000},
		{LayoutScreen16x9, 12192000},
		{LayoutScreen16x10, 10972800},
		{LayoutA4, 9906000},
		{LayoutLetter, 9144000},
	}

	for _, l := range layouts {
		dl := NewDocumentLayout()
		dl.SetLayout(l.name)
		if dl.CX != l.expectedCX {
			t.Errorf("layout %s: expected CX %d, got %d", l.name, l.expectedCX, dl.CX)
		}
		if dl.Name != l.name {
			t.Errorf("layout %s: expected name %s, got %s", l.name, l.name, dl.Name)
		}
	}

	// Custom layout
	dl := NewDocumentLayout()
	dl.SetCustomLayout(5000000, 4000000)
	if dl.CX != 5000000 || dl.CY != 4000000 {
		t.Error("custom layout dimensions mismatch")
	}
	if dl.Name != LayoutCustom {
		t.Errorf("expected custom layout name, got %s", dl.Name)
	}
}

// --- presentation.go: GetActiveSlide edge cases ---

func TestGetActiveSlideEdgeCases(t *testing.T) {
	// Empty presentation (no slides)
	p := &Presentation{
		properties:             NewDocumentProperties(),
		presentationProperties: NewPresentationProperties(),
		slides:                 make([]*Slide, 0),
		slideMasters:           make([]*SlideMaster, 0),
		layout:                 NewDocumentLayout(),
	}
	if p.GetActiveSlide() != nil {
		t.Error("expected nil for empty presentation")
	}

	// activeSlideIndex out of range
	p.CreateSlide()
	p.activeSlideIndex = 999
	slide := p.GetActiveSlide()
	if slide == nil {
		t.Fatal("expected non-nil slide after index reset")
	}
	if p.activeSlideIndex != 0 {
		t.Errorf("expected activeSlideIndex reset to 0, got %d", p.activeSlideIndex)
	}
}

// --- writer_slide.go: boolToWrap(false), writeBorderXML with border ---

func TestBoolToWrap(t *testing.T) {
	if boolToWrap(true) != "square" {
		t.Error("expected 'square' for true")
	}
	if boolToWrap(false) != "none" {
		t.Error("expected 'none' for false")
	}
}

func TestWriteBorderXMLWithBorder(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetHeight(200).SetWidth(400).SetOffsetX(0).SetOffsetY(0)
	rt.CreateTextRun("Bordered text")
	rt.SetBorder(&Border{Style: BorderSolid, Width: 2, Color: ColorRed})

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)
	if !strings.Contains(content, "a:ln") {
		t.Error("expected a:ln element for border")
	}
	if !strings.Contains(content, "FF0000") {
		t.Error("expected red color in border")
	}
}

func TestWriteWordWrapFalse(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetHeight(200).SetWidth(400).SetOffsetX(0).SetOffsetY(0)
	rt.SetWordWrap(false)
	rt.CreateTextRun("No wrap text")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)
	if !strings.Contains(content, `wrap="none"`) {
		t.Error("expected wrap='none' for wordWrap=false")
	}
}

// --- writer_chart.go: boolToXML, axisOrientation reversed ---

func TestBoolToXML(t *testing.T) {
	if boolToXML(true) != "1" {
		t.Error("expected '1' for true")
	}
	if boolToXML(false) != "0" {
		t.Error("expected '0' for false")
	}
}

func TestAxisOrientationReversed(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(100).SetOffsetY(100).SetWidth(5000000).SetHeight(3000000)

	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("Data", []string{"A", "B"}, []float64{10, 20}))
	chart.GetPlotArea().SetType(bar)
	chart.GetPlotArea().GetAxisX().SetReversedOrder(true)
	chart.GetPlotArea().GetAxisY().SetReversedOrder(true)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(data)
	if !strings.Contains(content, `orientation val="maxMin"`) {
		t.Error("expected maxMin orientation for reversed axis")
	}
}

// --- chart.go: SetBarGrouping stacked/percentStacked, SetOverlapPercent clamping ---

func TestBarChartGroupingAndOverlap(t *testing.T) {
	bar := NewBarChart()

	// Stacked sets overlap to 100
	bar.SetBarGrouping(BarGroupingStacked)
	if bar.OverlapPercent != 100 {
		t.Errorf("expected overlap 100 for stacked, got %d", bar.OverlapPercent)
	}

	// PercentStacked sets overlap to 100
	bar.SetBarGrouping(BarGroupingPercentStacked)
	if bar.OverlapPercent != 100 {
		t.Errorf("expected overlap 100 for percentStacked, got %d", bar.OverlapPercent)
	}

	// Clustered resets overlap to 0
	bar.SetBarGrouping(BarGroupingClustered)
	if bar.OverlapPercent != 0 {
		t.Errorf("expected overlap 0 for clustered, got %d", bar.OverlapPercent)
	}

	// SetOverlapPercent clamping
	bar.SetOverlapPercent(-200)
	if bar.OverlapPercent != -100 {
		t.Errorf("expected -100, got %d", bar.OverlapPercent)
	}
	bar.SetOverlapPercent(200)
	if bar.OverlapPercent != 100 {
		t.Errorf("expected 100, got %d", bar.OverlapPercent)
	}

	// SetGapWidthPercent clamping
	bar.SetGapWidthPercent(-10)
	if bar.GapWidthPercent != 0 {
		t.Errorf("expected 0, got %d", bar.GapWidthPercent)
	}
	bar.SetGapWidthPercent(600)
	if bar.GapWidthPercent != 500 {
		t.Errorf("expected 500, got %d", bar.GapWidthPercent)
	}
}

// --- shape.go: GetActiveParagraph edge case, TableCell.SetText ---

func TestRichTextGetActiveParagraphEmpty(t *testing.T) {
	rt := &RichTextShape{
		paragraphs: make([]*Paragraph, 0),
	}
	// Should auto-create a paragraph when empty
	para := rt.GetActiveParagraph()
	if para == nil {
		t.Fatal("expected non-nil paragraph")
	}
	if len(rt.paragraphs) != 1 {
		t.Errorf("expected 1 paragraph, got %d", len(rt.paragraphs))
	}
}

func TestTableCellSetTextMultipleCalls(t *testing.T) {
	cell := NewTableCell()
	cell.SetText("First")
	cell.SetText("Second")

	paras := cell.GetParagraphs()
	if len(paras) == 0 {
		t.Fatal("expected at least 1 paragraph")
	}
	// Both text runs should be in the first paragraph
	elems := paras[0].GetElements()
	if len(elems) < 2 {
		t.Fatalf("expected at least 2 elements, got %d", len(elems))
	}
}

// --- writer_slide.go: writeMedia with image data (various types) ---

func TestWriteMediaWithImageData(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// PNG image
	img1 := NewDrawingShape()
	img1.SetImageData([]byte{0x89, 0x50, 0x4E, 0x47}, "image/png")
	img1.SetWidth(100).SetHeight(100).SetOffsetX(0).SetOffsetY(0)
	slide.AddShape(img1)

	// JPEG image
	img2 := NewDrawingShape()
	img2.SetImageData([]byte{0xFF, 0xD8, 0xFF}, "image/jpeg")
	img2.SetWidth(100).SetHeight(100).SetOffsetX(200).SetOffsetY(0)
	slide.AddShape(img2)

	// GIF image
	img3 := NewDrawingShape()
	img3.SetImageData([]byte{0x47, 0x49, 0x46}, "image/gif")
	img3.SetWidth(100).SetHeight(100).SetOffsetX(400).SetOffsetY(0)
	slide.AddShape(img3)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Verify images exist
	for i := 1; i <= 3; i++ {
		var ext string
		switch i {
		case 1:
			ext = "png"
		case 2:
			ext = "jpeg"
		case 3:
			ext = "gif"
		}
		_, err := readFileFromZip(zr, fmt.Sprintf("ppt/media/image%d.%s", i, ext))
		if err != nil {
			t.Errorf("image%d.%s not found", i, ext)
		}
	}

	// Verify content types include all image types
	ctData, _ := readFileFromZip(zr, "[Content_Types].xml")
	ct := string(ctData)
	if !strings.Contains(ct, "image/png") {
		t.Error("content types should include image/png")
	}
	if !strings.Contains(ct, "image/jpeg") {
		t.Error("content types should include image/jpeg")
	}
	if !strings.Contains(ct, "image/gif") {
		t.Error("content types should include image/gif")
	}
}

// --- writer.go: Save() and WriteTo() error paths, NewWriter unsupported ---

func TestNewWriterUnsupportedFormat(t *testing.T) {
	_, err := NewWriter(New(), "unsupported")
	if err == nil {
		t.Error("expected error for unsupported format")
	}
}

func TestNewReaderUnsupportedFormat(t *testing.T) {
	_, err := NewReader("unsupported")
	if err == nil {
		t.Error("expected error for unsupported format")
	}
}

// --- writer_chart.go: getChartSeries for all chart types ---

func TestGetChartSeriesAllTypes(t *testing.T) {
	types := []struct {
		name string
		ct   ChartType
	}{
		{"bar", NewBarChart()},
		{"bar3D", NewBar3DChart()},
		{"line", NewLineChart()},
		{"area", NewAreaChart()},
		{"pie", NewPieChart()},
		{"pie3D", NewPie3DChart()},
		{"doughnut", NewDoughnutChart()},
		{"scatter", NewScatterChart()},
		{"radar", NewRadarChart()},
	}

	for _, tt := range types {
		series := getChartSeries(tt.ct)
		if series == nil {
			t.Errorf("getChartSeries(%s) returned nil", tt.name)
		}
		if len(series) != 0 {
			t.Errorf("getChartSeries(%s) expected 0 series, got %d", tt.name, len(series))
		}
	}

	// nil chart type
	series := getChartSeries(nil)
	if series != nil {
		t.Error("expected nil for nil chart type")
	}
}

// --- reader.go: readFileFromZip not found, readRelationships missing file ---

func TestReadFileFromZipNotFound(t *testing.T) {
	// Create a minimal zip
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	fw, _ := zw.Create("test.txt")
	fw.Write([]byte("hello"))
	zw.Close()

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	_, err := readFileFromZip(zr, "nonexistent.xml")
	if err == nil {
		t.Error("expected error for nonexistent file")
	}
}

func TestReadRelationshipsMissingFile(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	fw, _ := zw.Create("dummy.txt")
	fw.Write([]byte("x"))
	zw.Close()

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	reader := &PPTXReader{}
	rels, err := reader.readRelationships(zr, "nonexistent.rels")
	if err != nil {
		t.Errorf("expected nil error for missing rels, got %v", err)
	}
	if rels != nil {
		t.Error("expected nil rels for missing file")
	}
}

// --- writer_xml.go: writeXMLToZip error path ---

func TestWriteXMLToZipWellFormed(t *testing.T) {
	// Test that writeXMLToZip produces valid XML
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)

	type testXML struct {
		XMLName xml.Name `xml:"test"`
		Value   string   `xml:"value"`
	}
	err := writeXMLToZip(zw, "test.xml", testXML{Value: "hello"})
	if err != nil {
		t.Fatalf("writeXMLToZip error: %v", err)
	}
	zw.Close()

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, err := readFileFromZip(zr, "test.xml")
	if err != nil {
		t.Fatal("test.xml not found")
	}
	if !strings.Contains(string(data), "<test>") {
		t.Error("expected <test> element")
	}
	if !strings.Contains(string(data), "hello") {
		t.Error("expected 'hello' value")
	}
}

// --- Presentation: zoom in viewProps ---

func TestViewPropsZoom(t *testing.T) {
	p := New()
	p.GetPresentationProperties().SetZoom(2.5)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/viewProps.xml")
	content := string(data)
	// 2.5 * 100 = 250
	if !strings.Contains(content, `n="250"`) {
		t.Error("expected zoom scale n=250 in viewProps")
	}
}

// --- Empty presentation write (no shapes) ---

func TestWriteEmptyPresentation(t *testing.T) {
	p := New()
	// Slide exists but has no shapes

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	// Should still produce valid zip
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("invalid zip: %v", err)
	}

	// Verify essential files exist
	essentials := []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"ppt/presentation.xml",
		"ppt/slides/slide1.xml",
	}
	for _, name := range essentials {
		_, err := readFileFromZip(zr, name)
		if err != nil {
			t.Errorf("essential file %s not found", name)
		}
	}
}

// --- Chart: DisplayBlankAs modes ---

func TestChartDisplayBlankAs(t *testing.T) {
	chart := NewChartShape()
	if chart.GetDisplayBlankAs() != ChartBlankAsZero {
		t.Error("expected default ChartBlankAsZero")
	}

	chart.SetDisplayBlankAs(ChartBlankAsGap)
	if chart.GetDisplayBlankAs() != ChartBlankAsGap {
		t.Error("expected ChartBlankAsGap")
	}

	chart.SetDisplayBlankAs(ChartBlankAsSpan)
	if chart.GetDisplayBlankAs() != ChartBlankAsSpan {
		t.Error("expected ChartBlankAsSpan")
	}
}

// --- Chart title visibility: hidden title ---

func TestChartHiddenTitle(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)
	chart.GetTitle().SetVisible(false)

	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("S", []string{"A"}, []float64{1}))
	chart.GetPlotArea().SetType(bar)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(data)
	if !strings.Contains(content, "autoTitleDeleted") {
		t.Error("expected autoTitleDeleted for hidden title")
	}
}

// --- Chart with bold title ---

func TestChartBoldTitle(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)
	chart.GetTitle().SetText("Bold Title").SetVisible(true)
	chart.GetTitle().Font.SetBold(true)

	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("S", []string{"A"}, []float64{1}))
	chart.GetPlotArea().SetType(bar)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(data)
	if !strings.Contains(content, `b="1"`) {
		t.Error("expected bold attribute in chart title")
	}
}

// --- isPieType helper ---

func TestIsPieType(t *testing.T) {
	if !isPieType(NewPieChart()) {
		t.Error("PieChart should be pie type")
	}
	if !isPieType(NewPie3DChart()) {
		t.Error("Pie3DChart should be pie type")
	}
	if !isPieType(NewDoughnutChart()) {
		t.Error("DoughnutChart should be pie type")
	}
	if isPieType(NewBarChart()) {
		t.Error("BarChart should not be pie type")
	}
	if isPieType(NewLineChart()) {
		t.Error("LineChart should not be pie type")
	}
}

// ===== Coverage Expansion Phase 2 =====

// --- writer_slide.go: writeParagraphXML alignment level, spacing ---

func TestWriteParagraphWithAlignmentLevel(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetHeight(500).SetWidth(800).SetOffsetX(0).SetOffsetY(0)

	para := rt.GetActiveParagraph()
	para.GetAlignment().SetHorizontal(HorizontalCenter)
	para.GetAlignment().Level = 2
	para.SetLineSpacing(300)
	para.SetSpaceBefore(200)
	para.SetSpaceAfter(100)
	para.CreateTextRun("Indented centered text")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	if !strings.Contains(content, `algn="ctr"`) {
		t.Error("expected center alignment")
	}
	if !strings.Contains(content, `lvl="2"`) {
		t.Error("expected level 2")
	}
	if !strings.Contains(content, "a:lnSpc") {
		t.Error("expected line spacing element")
	}
	if !strings.Contains(content, "a:spcBef") {
		t.Error("expected space before element")
	}
	if !strings.Contains(content, "a:spcAft") {
		t.Error("expected space after element")
	}
}

// --- writer_slide.go: writeTextRunXML strikethrough, underline, hyperlink ---

func TestWriteTextRunFormatting(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetHeight(500).SetWidth(800).SetOffsetX(0).SetOffsetY(0)

	// Strikethrough text
	tr1 := rt.GetActiveParagraph().CreateTextRun("Struck text")
	tr1.GetFont().SetStrikethrough(true)

	// Underline text
	para2 := rt.CreateParagraph()
	tr2 := para2.CreateTextRun("Underlined text")
	tr2.GetFont().SetUnderline(UnderlineSingle)

	// Bold italic with color
	para3 := rt.CreateParagraph()
	tr3 := para3.CreateTextRun("Bold italic")
	tr3.GetFont().SetBold(true).SetItalic(true).SetColor(ColorRed).SetName("Arial")

	// Hyperlink text
	para4 := rt.CreateParagraph()
	tr4 := para4.CreateTextRun("Click here")
	tr4.SetHyperlink(NewHyperlink("https://example.com"))

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	if !strings.Contains(content, `strike="sngStrike"`) {
		t.Error("expected strikethrough attribute")
	}
	if !strings.Contains(content, `u="sng"`) {
		t.Error("expected underline attribute")
	}
	if !strings.Contains(content, `b="1"`) {
		t.Error("expected bold attribute")
	}
	if !strings.Contains(content, `i="1"`) {
		t.Error("expected italic attribute")
	}
	if !strings.Contains(content, "FF0000") {
		t.Error("expected red color")
	}
	if !strings.Contains(content, `typeface="Arial"`) {
		t.Error("expected Arial typeface")
	}
	if !strings.Contains(content, "hlinkClick") {
		t.Error("expected hyperlink element")
	}

	// Verify hyperlink in rels
	relData, _ := readFileFromZip(zr, "ppt/slides/_rels/slide1.xml.rels")
	relContent := string(relData)
	if !strings.Contains(relContent, "https://example.com") {
		t.Error("expected hyperlink URL in rels")
	}
	if !strings.Contains(relContent, "External") {
		t.Error("expected External target mode for hyperlink")
	}
}

// --- writer_slide.go: writeDrawingShapeXML with shadow ---

func TestWriteDrawingWithShadow(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	img := NewDrawingShape()
	img.SetImageData([]byte{0x89, 0x50, 0x4E, 0x47}, "image/png")
	img.SetWidth(1000000).SetHeight(800000).SetOffsetX(100).SetOffsetY(100)
	img.SetDescription("Test image")
	img.SetShadow(&Shadow{
		Visible:    true,
		Direction:  45,
		Distance:   5,
		BlurRadius: 3,
		Color:      Color{ARGB: "FF000000"},
		Alpha:      50,
	})
	slide.AddShape(img)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	if !strings.Contains(content, "a:outerShdw") {
		t.Error("expected outer shadow element")
	}
	if !strings.Contains(content, "a:effectLst") {
		t.Error("expected effect list element")
	}
	if !strings.Contains(content, "a:alpha") {
		t.Error("expected alpha element in shadow")
	}
}

// --- writer_slide.go: writeGroupShapeXML with AutoShape and DrawingShape children ---

func TestWriteGroupShapeWithVariousChildren(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	group := slide.CreateGroupShape()
	group.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)

	// AutoShape child
	as := NewAutoShape()
	as.SetAutoShapeType(AutoShapeEllipse)
	as.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(1000000).SetHeight(1000000)
	group.AddShape(as)

	// Drawing child
	img := NewDrawingShape()
	img.SetImageData([]byte{0x89, 0x50}, "image/png")
	img.SetWidth(500000).SetHeight(500000).SetOffsetX(1000000).SetOffsetY(0)
	group.AddShape(img)

	// Line child
	line := NewLineShape()
	line.BaseShape.SetOffsetX(0).SetOffsetY(1000000).SetWidth(2000000).SetHeight(0)
	group.AddShape(line)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	if !strings.Contains(content, "grpSp") {
		t.Error("expected grpSp element")
	}
	if !strings.Contains(content, "ellipse") {
		t.Error("expected ellipse auto shape in group")
	}
	if !strings.Contains(content, "p:pic") {
		t.Error("expected pic element in group")
	}
	if !strings.Contains(content, "p:cxnSp") {
		t.Error("expected cxnSp (line) element in group")
	}
}

// --- writer_slide.go: writeFillXML gradient path ---

func TestWriteGradientFillOnShape(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetHeight(200).SetWidth(400).SetOffsetX(0).SetOffsetY(0)
	rt.CreateTextRun("Gradient fill")
	rt.SetFill(NewFill().SetGradientLinear(ColorRed, ColorBlue, 90))

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	if !strings.Contains(content, "a:gradFill") {
		t.Error("expected gradient fill element")
	}
	if !strings.Contains(content, "a:gsLst") {
		t.Error("expected gradient stop list")
	}
	if !strings.Contains(content, "a:lin") {
		t.Error("expected linear gradient element")
	}
}

// --- writer_chart.go: writeSeriesXML data labels, markers, fill, separator ---

func TestWriteChartSeriesWithDataLabels(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)

	bar := NewBarChart()
	s := NewChartSeriesOrdered("Sales", []string{"Q1", "Q2"}, []float64{100, 200})
	s.ShowValue = true
	s.ShowCategoryName = true
	s.ShowSeriesName = true
	s.ShowPercentage = true
	s.Separator = " | "
	s.LabelPosition = LabelOutsideEnd
	s.SetFillColor(ColorRed)
	bar.AddSeries(s)
	chart.GetPlotArea().SetType(bar)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(data)

	if !strings.Contains(content, "c:dLbls") {
		t.Error("expected data labels element")
	}
	if !strings.Contains(content, `showVal val="1"`) {
		t.Error("expected showVal")
	}
	if !strings.Contains(content, `showCatName val="1"`) {
		t.Error("expected showCatName")
	}
	if !strings.Contains(content, `showSerName val="1"`) {
		t.Error("expected showSerName")
	}
	if !strings.Contains(content, `showPercent val="1"`) {
		t.Error("expected showPercent")
	}
	if !strings.Contains(content, "c:separator") {
		t.Error("expected separator element")
	}
	if !strings.Contains(content, `dLblPos val="outEnd"`) {
		t.Error("expected label position")
	}
	// Fill color on series
	if !strings.Contains(content, "FF0000") {
		t.Error("expected red fill color on series")
	}
}

func TestWriteLineChartWithMarkers(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)

	line := NewLineChart()
	line.SetSmooth(true)
	s := NewChartSeriesOrdered("Trend", []string{"Jan", "Feb", "Mar"}, []float64{10, 25, 15})
	s.Marker = &SeriesMarker{Symbol: MarkerCircle, Size: 5}
	line.AddSeries(s)
	chart.GetPlotArea().SetType(line)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(data)

	if !strings.Contains(content, "c:lineChart") {
		t.Error("expected lineChart element")
	}
	if !strings.Contains(content, `c:smooth val="1"`) {
		t.Error("expected smooth=1")
	}
	if !strings.Contains(content, "c:marker") {
		t.Error("expected marker element")
	}
	if !strings.Contains(content, `c:symbol val="circle"`) {
		t.Error("expected circle marker symbol")
	}
}

// --- writer_chart.go: writeAxesXML minor gridlines, units ---

func TestWriteChartAxesWithGridlinesAndUnits(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)

	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("D", []string{"A", "B"}, []float64{10, 20}))
	chart.GetPlotArea().SetType(bar)

	// X axis with major gridlines
	chart.GetPlotArea().GetAxisX().SetMajorGridlines(NewGridlines())

	// Y axis with bounds, units, and both gridlines
	axY := chart.GetPlotArea().GetAxisY()
	axY.SetMinBounds(0).SetMaxBounds(100)
	axY.SetMajorUnit(20).SetMinorUnit(5)
	axY.SetMajorGridlines(NewGridlines())
	axY.SetMinorGridlines(&Gridlines{Width: 1, Color: Color{ARGB: "FF888888"}})
	axY.SetTitle("Values")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(data)

	if !strings.Contains(content, "c:majorGridlines") {
		t.Error("expected major gridlines")
	}
	if !strings.Contains(content, "c:minorGridlines") {
		t.Error("expected minor gridlines")
	}
	if !strings.Contains(content, `c:majorUnit val="20"`) {
		t.Error("expected major unit 20")
	}
	if !strings.Contains(content, `c:minorUnit val="5"`) {
		t.Error("expected minor unit 5")
	}
	if !strings.Contains(content, `c:min val="0"`) {
		t.Error("expected min bound 0")
	}
	if !strings.Contains(content, `c:max val="100"`) {
		t.Error("expected max bound 100")
	}
}

// --- writer_chart.go: scatter chart with fill color ---

func TestWriteScatterChartWithFillColor(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)

	scatter := NewScatterChart()
	scatter.SetSmooth(true)
	s := NewChartSeriesOrdered("Points", []string{"1", "2", "3"}, []float64{10, 20, 30})
	s.SetFillColor(ColorGreen)
	scatter.AddSeries(s)
	chart.GetPlotArea().SetType(scatter)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(data)

	if !strings.Contains(content, "c:scatterChart") {
		t.Error("expected scatterChart element")
	}
	if !strings.Contains(content, "00FF00") {
		t.Error("expected green fill color on scatter series")
	}
	if !strings.Contains(content, `c:smooth val="1"`) {
		t.Error("expected smooth=1 on scatter")
	}
}

// --- writer_chart.go: getCategories with empty series ---

func TestGetCategoriesEmpty(t *testing.T) {
	cats := getCategories(nil)
	if cats != nil {
		t.Error("expected nil for nil series")
	}
	cats = getCategories([]*ChartSeries{})
	if cats != nil {
		t.Error("expected nil for empty series")
	}
}

// --- writer_comment.go: collectAuthors duplicate author ---

func TestCollectAuthorsDuplicate(t *testing.T) {
	p := New()
	slide1 := p.GetActiveSlide()
	slide2 := p.CreateSlide()

	author := NewCommentAuthor("Same Author", "SA")
	slide1.AddComment(NewComment().SetAuthor(author).SetText("Comment 1"))
	slide2.AddComment(NewComment().SetAuthor(author).SetText("Comment 2"))

	// Also add a comment with nil author
	slide1.AddComment(NewComment().SetText("No author comment"))

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	authData, _ := readFileFromZip(zr, "ppt/commentAuthors.xml")
	content := string(authData)

	// Should only have one author entry despite two comments
	// p:cmAuthorLst contains "p:cmAuthor" as substring, so count actual author elements
	authorCount := strings.Count(content, `<p:cmAuthor `)
	if authorCount != 1 {
		t.Errorf("expected 1 author entry, got %d", authorCount)
	}
}

// --- shape.go: TableCell.SetText with empty paragraphs ---

func TestTableCellSetTextEmptyParagraphs(t *testing.T) {
	cell := &TableCell{
		paragraphs: make([]*Paragraph, 0),
		fill:       NewFill(),
		border: &CellBorders{
			Top: NewBorder(), Bottom: NewBorder(),
			Left: NewBorder(), Right: NewBorder(),
		},
	}
	cell.SetText("Hello")
	if len(cell.paragraphs) != 1 {
		t.Errorf("expected 1 paragraph, got %d", len(cell.paragraphs))
	}
	elems := cell.paragraphs[0].GetElements()
	if len(elems) != 1 {
		t.Fatalf("expected 1 element, got %d", len(elems))
	}
	if tr, ok := elems[0].(*TextRun); ok {
		if tr.GetText() != "Hello" {
			t.Errorf("expected 'Hello', got '%s'", tr.GetText())
		}
	}
}

// --- writer_slide.go: RichTextShape with rotation ---

func TestWriteRichTextWithRotation(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetHeight(200).SetWidth(400).SetOffsetX(0).SetOffsetY(0)
	rt.BaseShape.SetRotation(45)
	rt.CreateTextRun("Rotated text")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	// 45 degrees * 60000 = 2700000
	if !strings.Contains(content, `rot="2700000"`) {
		t.Error("expected rotation attribute")
	}
}

// --- writer_slide.go: writeChartShapeXML with multiple charts and images ---

func TestWriteSlideWithMultipleChartsAndImages(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// Image first
	img := NewDrawingShape()
	img.SetImageData([]byte{0x89, 0x50}, "image/png")
	img.SetWidth(100).SetHeight(100).SetOffsetX(0).SetOffsetY(0)
	slide.AddShape(img)

	// First chart
	chart1 := slide.CreateChartShape()
	chart1.BaseShape.SetOffsetX(200).SetOffsetY(0).SetWidth(3000000).SetHeight(2000000)
	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("S1", []string{"A"}, []float64{10}))
	chart1.GetPlotArea().SetType(bar)

	// Second chart
	chart2 := slide.CreateChartShape()
	chart2.BaseShape.SetOffsetX(200).SetOffsetY(2500000).SetWidth(3000000).SetHeight(2000000)
	pie := NewPieChart()
	pie.AddSeries(NewChartSeriesOrdered("S2", []string{"X", "Y"}, []float64{60, 40}))
	chart2.GetPlotArea().SetType(pie)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Both charts should exist
	_, err = readFileFromZip(zr, "ppt/charts/chart1.xml")
	if err != nil {
		t.Error("chart1.xml not found")
	}
	_, err = readFileFromZip(zr, "ppt/charts/chart2.xml")
	if err != nil {
		t.Error("chart2.xml not found")
	}

	// Image should exist
	_, err = readFileFromZip(zr, "ppt/media/image1.png")
	if err != nil {
		t.Error("image1.png not found")
	}

	// Slide rels should reference both charts and image
	relData, _ := readFileFromZip(zr, "ppt/slides/_rels/slide1.xml.rels")
	relContent := string(relData)
	if !strings.Contains(relContent, "chart1.xml") {
		t.Error("rels should reference chart1")
	}
	if !strings.Contains(relContent, "chart2.xml") {
		t.Error("rels should reference chart2")
	}
	if !strings.Contains(relContent, "image1.png") {
		t.Error("rels should reference image")
	}
}

// --- writer_slide.go: writeNotesSlide error path (already tested success) ---
// --- writer_slide.go: writeBulletXML buNone path ---

func TestWriteBulletNone(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetHeight(200).SetWidth(400).SetOffsetX(0).SetOffsetY(0)

	para := rt.GetActiveParagraph()
	para.SetBullet(&Bullet{Type: BulletTypeNone})
	para.CreateTextRun("No bullet")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	if !strings.Contains(content, "a:buNone") {
		t.Error("expected buNone element")
	}
}

// --- writer_chart.go: writeChartPart with nil chart type ---

func TestWriteChartPartNilType(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// Chart with no type set
	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(3000000).SetHeight(2000000)
	// Don't set chart type

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	// Should still produce valid output (chart part skipped)
	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	_, err = readFileFromZip(zr, "ppt/slides/slide1.xml")
	if err != nil {
		t.Error("slide1.xml not found")
	}
}

// --- writer_slide.go: AutoShape with text ---

func TestWriteAutoShapeWithText(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	as := slide.CreateAutoShape()
	as.SetAutoShapeType(AutoShapeRoundedRect)
	as.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(2000000).SetHeight(1000000)
	as.SetText("Inside shape")
	as.BaseShape.SetFill(NewFill().SetSolid(ColorYellow))

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	if !strings.Contains(content, "roundRect") {
		t.Error("expected roundRect preset geometry")
	}
	if !strings.Contains(content, "Inside shape") {
		t.Error("expected text in auto shape")
	}
	if !strings.Contains(content, "FFFF00") {
		t.Error("expected yellow fill")
	}
}

// --- Read round-trip: table with cell fill ---

func TestWriteReadTableWithCellFill(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	table := slide.CreateTableShape(2, 2)
	table.SetWidth(4000000).SetHeight(1000000)
	table.BaseShape.SetOffsetX(100).SetOffsetY(100)

	cell := table.GetCell(0, 0)
	cell.SetText("Header")
	cell.SetFill(NewFill().SetSolid(ColorBlue))

	table.GetCell(0, 1).SetText("Value")
	table.GetCell(1, 0).SetText("Row2")
	table.GetCell(1, 1).SetText("Data")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	if !strings.Contains(content, "a:tbl") {
		t.Error("expected table element")
	}
	if !strings.Contains(content, "0000FF") {
		t.Error("expected blue cell fill")
	}
	if !strings.Contains(content, "Header") {
		t.Error("expected Header text")
	}

	// Read back
	reader := &PPTXReader{}
	pres, err := reader.ReadFromReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	shapes := pres.GetActiveSlide().GetShapes()
	var tbl *TableShape
	for _, s := range shapes {
		if tb, ok := s.(*TableShape); ok {
			tbl = tb
			break
		}
	}
	if tbl == nil {
		t.Fatal("expected TableShape")
	}
	if tbl.GetNumRows() != 2 || tbl.GetNumCols() != 2 {
		t.Errorf("expected 2x2 table, got %dx%d", tbl.GetNumRows(), tbl.GetNumCols())
	}
}

// --- Read round-trip: line shape ---

func TestWriteReadLineShape(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	line := slide.CreateLineShape()
	line.BaseShape.SetOffsetX(100).SetOffsetY(200).SetWidth(3000000).SetHeight(0)
	line.SetLineWidth(3).SetLineColor(ColorRed)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	reader := &PPTXReader{}
	pres, err := reader.ReadFromReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	shapes := pres.GetActiveSlide().GetShapes()
	var readLine *LineShape
	for _, s := range shapes {
		if l, ok := s.(*LineShape); ok {
			readLine = l
			break
		}
	}
	if readLine == nil {
		t.Fatal("expected LineShape")
	}
	if readLine.GetWidth() != 3000000 {
		t.Errorf("expected width 3000000, got %d", readLine.GetWidth())
	}
}

// --- Read round-trip: image shape ---

func TestWriteReadImageShape(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	img := NewDrawingShape()
	img.SetImageData([]byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A}, "image/png")
	img.SetWidth(2000000).SetHeight(1500000).SetOffsetX(500).SetOffsetY(500)
	img.SetName("TestImage")
	slide.AddShape(img)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	reader := &PPTXReader{}
	pres, err := reader.ReadFromReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	shapes := pres.GetActiveSlide().GetShapes()
	var readImg *DrawingShape
	for _, s := range shapes {
		if d, ok := s.(*DrawingShape); ok {
			readImg = d
			break
		}
	}
	if readImg == nil {
		t.Fatal("expected DrawingShape")
	}
	if readImg.GetWidth() != 2000000 {
		t.Errorf("expected width 2000000, got %d", readImg.GetWidth())
	}
	if len(readImg.GetImageData()) == 0 {
		t.Error("expected image data to be read back")
	}
	if readImg.GetMimeType() != "image/png" {
		t.Errorf("expected image/png, got %s", readImg.GetMimeType())
	}
}

// --- Read: multiple slides with various shapes ---

func TestReadMultiSlidePresentation(t *testing.T) {
	p := New()

	// Slide 1: rich text
	s1 := p.GetActiveSlide()
	rt := s1.CreateRichTextShape()
	rt.SetHeight(200).SetWidth(400).SetOffsetX(0).SetOffsetY(0)
	rt.CreateTextRun("Slide 1")

	// Slide 2: auto shape
	s2 := p.CreateSlide()
	as := s2.CreateAutoShape()
	as.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(1000000).SetHeight(1000000)

	// Slide 3: line
	s3 := p.CreateSlide()
	line := s3.CreateLineShape()
	line.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(2000000).SetHeight(0)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	reader := &PPTXReader{}
	pres, err := reader.ReadFromReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	if pres.GetSlideCount() != 3 {
		t.Errorf("expected 3 slides, got %d", pres.GetSlideCount())
	}
}

// --- writer_slide.go: break element in paragraph ---

func TestWriteBreakElement(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetHeight(300).SetWidth(500).SetOffsetX(0).SetOffsetY(0)

	para := rt.GetActiveParagraph()
	para.CreateTextRun("Line 1")
	para.CreateBreak()
	para.CreateTextRun("Line 2")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	if !strings.Contains(content, "a:br") {
		t.Error("expected break element")
	}
	if !strings.Contains(content, "Line 1") {
		t.Error("expected Line 1")
	}
	if !strings.Contains(content, "Line 2") {
		t.Error("expected Line 2")
	}
}

// --- writer.go: Save to temp file ---

func TestSaveToTempFile(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	rt := slide.CreateRichTextShape()
	rt.SetHeight(100).SetWidth(200).SetOffsetX(0).SetOffsetY(0)
	rt.CreateTextRun("Save test")

	tmpDir := t.TempDir()
	path := tmpDir + "/test_save.pptx"

	w, _ := NewWriter(p, WriterPowerPoint2007)
	pptxWriter := w.(*PPTXWriter)
	err := pptxWriter.Save(path)
	if err != nil {
		t.Fatalf("Save error: %v", err)
	}

	// Read back from file
	reader := &PPTXReader{}
	pres, err := reader.Read(path)
	if err != nil {
		t.Fatalf("Read error: %v", err)
	}
	if pres.GetSlideCount() != 1 {
		t.Errorf("expected 1 slide, got %d", pres.GetSlideCount())
	}
}

// --- writer.go: Save to nested directory ---

func TestSaveToNestedDirectory(t *testing.T) {
	p := New()
	p.GetActiveSlide().CreateRichTextShape().SetHeight(100).SetWidth(200)

	tmpDir := t.TempDir()
	path := tmpDir + "/sub/dir/test.pptx"

	w, _ := NewWriter(p, WriterPowerPoint2007)
	pptxWriter := w.(*PPTXWriter)
	err := pptxWriter.Save(path)
	if err != nil {
		t.Fatalf("Save error: %v", err)
	}

	reader := &PPTXReader{}
	pres, err := reader.Read(path)
	if err != nil {
		t.Fatalf("Read error: %v", err)
	}
	if pres.GetSlideCount() != 1 {
		t.Errorf("expected 1 slide, got %d", pres.GetSlideCount())
	}
}

// --- Doughnut chart write ---

func TestWriteDoughnutChart(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)

	d := NewDoughnutChart()
	d.HoleSize = 75
	d.AddSeries(NewChartSeriesOrdered("Parts", []string{"A", "B", "C"}, []float64{30, 50, 20}))
	chart.GetPlotArea().SetType(d)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(data)

	if !strings.Contains(content, "c:doughnutChart") {
		t.Error("expected doughnutChart element")
	}
	if !strings.Contains(content, `holeSize val="75"`) {
		t.Error("expected holeSize 75")
	}
}

// --- Area chart write ---

func TestWriteAreaChart(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)

	area := NewAreaChart()
	area.AddSeries(NewChartSeriesOrdered("Area", []string{"Jan", "Feb", "Mar"}, []float64{10, 30, 20}))
	chart.GetPlotArea().SetType(area)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(data)

	if !strings.Contains(content, "c:areaChart") {
		t.Error("expected areaChart element")
	}
}

// --- Radar chart write ---

func TestWriteRadarChart(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(0).SetOffsetY(0).SetWidth(5000000).SetHeight(3000000)

	radar := NewRadarChart()
	radar.AddSeries(NewChartSeriesOrdered("Skills", []string{"A", "B", "C", "D"}, []float64{80, 90, 70, 85}))
	chart.GetPlotArea().SetType(radar)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/charts/chart1.xml")
	content := string(data)

	if !strings.Contains(content, "c:radarChart") {
		t.Error("expected radarChart element")
	}
	if !strings.Contains(content, `radarStyle val="marker"`) {
		t.Error("expected radar style marker")
	}
}

// ===== Coverage Expansion Phase 3 =====

// --- writer_slide.go: writeMedia with file-path image ---

func TestWriteMediaWithFilePath(t *testing.T) {
	// Create a temp image file
	tmpDir := t.TempDir()
	imgPath := tmpDir + "/test.png"
	// Write minimal PNG-like bytes
	imgData := []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A}
	if err := os.WriteFile(imgPath, imgData, 0644); err != nil {
		t.Fatalf("failed to write temp image: %v", err)
	}

	p := New()
	slide := p.GetActiveSlide()

	img := NewDrawingShape()
	img.SetPath(imgPath)
	img.SetWidth(1000000).SetHeight(800000).SetOffsetX(0).SetOffsetY(0)
	slide.AddShape(img)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	_, err = readFileFromZip(zr, "ppt/media/image1.png")
	if err != nil {
		t.Error("image1.png not found in zip")
	}
}

// --- writer_slide.go: writeMedia with nonexistent file path (error) ---

func TestWriteMediaWithBadFilePath(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	img := NewDrawingShape()
	img.SetPath("/nonexistent/path/image.png")
	img.SetWidth(100).SetHeight(100).SetOffsetX(0).SetOffsetY(0)
	slide.AddShape(img)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err == nil {
		t.Error("expected error for nonexistent image path")
	}
}

// --- writer_slide.go: writeFillXML with nil fill ---

func TestWriteFillXMLNilAndNone(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// Shape with no fill set (nil)
	rt := slide.CreateRichTextShape()
	rt.SetHeight(200).SetWidth(400).SetOffsetX(0).SetOffsetY(0)
	rt.CreateTextRun("No fill")
	// Explicitly set fill to nil
	rt.SetFill(nil)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	// Should produce valid output without fill elements
	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)
	if !strings.Contains(content, "No fill") {
		t.Error("expected text content")
	}
}

// --- writer_slide.go: getImageIndex across multiple slides ---

func TestGetImageIndexMultiSlides(t *testing.T) {
	p := New()

	// Slide 1 with image
	s1 := p.GetActiveSlide()
	img1 := NewDrawingShape()
	img1.SetImageData([]byte{0x89}, "image/png")
	img1.SetWidth(100).SetHeight(100).SetOffsetX(0).SetOffsetY(0)
	s1.AddShape(img1)

	// Slide 2 with image
	s2 := p.CreateSlide()
	img2 := NewDrawingShape()
	img2.SetImageData([]byte{0xFF, 0xD8}, "image/jpeg")
	img2.SetWidth(100).SetHeight(100).SetOffsetX(0).SetOffsetY(0)
	s2.AddShape(img2)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Both images should exist
	_, err = readFileFromZip(zr, "ppt/media/image1.png")
	if err != nil {
		t.Error("image1.png not found")
	}
	_, err = readFileFromZip(zr, "ppt/media/image2.jpeg")
	if err != nil {
		t.Error("image2.jpeg not found")
	}

	// Slide 2 rels should reference image2
	relData, _ := readFileFromZip(zr, "ppt/slides/_rels/slide2.xml.rels")
	if !strings.Contains(string(relData), "image2") {
		t.Error("slide2 rels should reference image2")
	}
}

// --- reader_parts.go: readPresentation fallback to rels ---
// The readPresentation function has a fallback path when sldId doesn't have
// namespace-qualified r:id. We test this by writing and reading back normally
// which exercises the fallback path.

// --- reader_slide.go: readSlide with all shape types in one slide ---

func TestReadSlideAllShapeTypes(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// RichText
	rt := slide.CreateRichTextShape()
	rt.SetHeight(200).SetWidth(400).SetOffsetX(0).SetOffsetY(0)
	rt.SetName("MyText")
	rt.CreateTextRun("Hello")

	// AutoShape
	as := slide.CreateAutoShape()
	as.BaseShape.SetOffsetX(500).SetOffsetY(0).SetWidth(1000000).SetHeight(1000000)

	// Line
	line := slide.CreateLineShape()
	line.BaseShape.SetOffsetX(0).SetOffsetY(1500000).SetWidth(2000000).SetHeight(0)

	// Table
	table := slide.CreateTableShape(1, 2)
	table.SetWidth(3000000).SetHeight(500000)
	table.BaseShape.SetOffsetX(0).SetOffsetY(2000000)
	table.GetCell(0, 0).SetText("Cell1")
	table.GetCell(0, 1).SetText("Cell2")

	// Image
	img := NewDrawingShape()
	img.SetImageData([]byte{0x89, 0x50, 0x4E, 0x47}, "image/png")
	img.SetWidth(500000).SetHeight(500000).SetOffsetX(4000000).SetOffsetY(0)
	slide.AddShape(img)

	// Placeholder
	ph := slide.CreatePlaceholderShape(PlaceholderBody)
	ph.BaseShape.SetOffsetX(0).SetOffsetY(3000000).SetWidth(8000000).SetHeight(500000)
	ph.CreateTextRun("Body text")

	// Group with child
	group := slide.CreateGroupShape()
	group.BaseShape.SetOffsetX(0).SetOffsetY(4000000).SetWidth(3000000).SetHeight(1000000)
	childRT := NewRichTextShape()
	childRT.SetOffsetX(0).SetOffsetY(0).SetWidth(1000000).SetHeight(500000)
	childRT.CreateTextRun("In group")
	group.AddShape(childRT)

	// Notes and background
	slide.SetNotes("Test notes")
	slide.SetBackground(NewFill().SetSolid(ColorWhite))

	// Comment
	author := NewCommentAuthor("Author", "A")
	slide.AddComment(NewComment().SetAuthor(author).SetText("Test comment").SetPosition(10, 20))

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	// Read back
	reader := &PPTXReader{}
	pres, err := reader.ReadFromReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("ReadFromReader error: %v", err)
	}

	readSlide := pres.GetActiveSlide()
	shapes := readSlide.GetShapes()

	// Count shape types
	var rtCount, lineCount, tableCount, imgCount, phCount, grpCount int
	for _, s := range shapes {
		switch s.(type) {
		case *RichTextShape:
			rtCount++
		case *LineShape:
			lineCount++
		case *TableShape:
			tableCount++
		case *DrawingShape:
			imgCount++
		case *PlaceholderShape:
			phCount++
		case *GroupShape:
			grpCount++
		}
	}

	if rtCount < 1 {
		t.Error("expected at least 1 RichTextShape")
	}
	if lineCount < 1 {
		t.Error("expected at least 1 LineShape")
	}
	if tableCount < 1 {
		t.Error("expected at least 1 TableShape")
	}
	if imgCount < 1 {
		t.Error("expected at least 1 DrawingShape")
	}
	if phCount < 1 {
		t.Error("expected at least 1 PlaceholderShape")
	}
	if grpCount < 1 {
		t.Error("expected at least 1 GroupShape")
	}

	// Notes
	if readSlide.GetNotes() != "Test notes" {
		t.Errorf("expected 'Test notes', got '%s'", readSlide.GetNotes())
	}

	// Background
	bg := readSlide.GetBackground()
	if bg == nil || bg.Type != FillSolid {
		t.Error("expected solid background")
	}

	// Comments
	if readSlide.GetCommentCount() != 1 {
		t.Errorf("expected 1 comment, got %d", readSlide.GetCommentCount())
	}
}

// --- reader.go: Read from invalid zip data ---

func TestReadFromInvalidData(t *testing.T) {
	reader := &PPTXReader{}
	_, err := reader.ReadFromReader(bytes.NewReader([]byte("not a zip")), 10)
	if err == nil {
		t.Error("expected error for invalid zip data")
	}
}

// --- writer_xml.go: getImageContentType default path ---

func TestGetImageContentTypeDefault(t *testing.T) {
	w := &PPTXWriter{presentation: New()}
	ds := &DrawingShape{}
	ds.path = "file.xyz" // unknown extension
	ct := w.getImageContentType(ds)
	if ct != "image/png" {
		t.Errorf("expected default image/png, got %s", ct)
	}
}
