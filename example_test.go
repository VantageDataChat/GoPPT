package gopresentation_test

import (
	"fmt"
	"os"

	gp "github.com/VantageDataChat/GoPPT"
)

func Example() {
	// Create a new presentation
	pres := gp.New()

	// Set document properties
	props := pres.GetDocumentProperties()
	props.Creator = "GoPresentation"
	props.Title = "Sample Presentation"
	props.Description = "Created with GoPresentation"

	// Set 16:9 layout
	pres.GetLayout().SetLayout(gp.LayoutScreen16x9)

	// --- Slide 1: Title slide ---
	slide1 := pres.GetActiveSlide()
	slide1.SetName("Title Slide")

	title := slide1.CreateRichTextShape()
	title.SetHeight(1000000).SetWidth(8000000).SetOffsetX(2000000).SetOffsetY(2000000)
	tr := title.CreateTextRun("Welcome to GoPresentation")
	tr.GetFont().SetBold(true).SetSize(36).SetColor(gp.NewColor("4472C4"))
	title.GetActiveParagraph().GetAlignment().SetHorizontal(gp.HorizontalCenter)

	subtitle := slide1.CreateRichTextShape()
	subtitle.SetHeight(500000).SetWidth(8000000).SetOffsetX(2000000).SetOffsetY(3500000)
	str := subtitle.CreateTextRun("A pure Go library for PowerPoint files")
	str.GetFont().SetSize(18).SetColor(gp.NewColor("666666"))
	subtitle.GetActiveParagraph().GetAlignment().SetHorizontal(gp.HorizontalCenter)

	// --- Slide 2: Shapes demo ---
	slide2 := pres.CreateSlide()
	slide2.SetName("Shapes Demo")

	// Rectangle with fill
	rect := slide2.CreateAutoShape()
	rect.SetAutoShapeType(gp.AutoShapeRectangle)
	rect.BaseShape.SetOffsetX(500000).SetOffsetY(500000).SetWidth(3000000).SetHeight(2000000)
	rect.GetFill().SetSolid(gp.NewColor("4472C4"))
	rect.SetText("Rectangle")

	// Ellipse
	ellipse := slide2.CreateAutoShape()
	ellipse.SetAutoShapeType(gp.AutoShapeEllipse)
	ellipse.BaseShape.SetOffsetX(4000000).SetOffsetY(500000).SetWidth(3000000).SetHeight(2000000)
	ellipse.GetFill().SetSolid(gp.NewColor("ED7D31"))

	// Line
	line := slide2.CreateLineShape()
	line.BaseShape.SetOffsetX(500000).SetOffsetY(3000000).SetWidth(7000000).SetHeight(0)
	line.SetLineWidth(2).SetLineColor(gp.NewColor("A5A5A5"))

	// --- Slide 3: Table ---
	slide3 := pres.CreateSlide()
	slide3.SetName("Table Demo")

	table := slide3.CreateTableShape(3, 3)
	table.SetWidth(8000000).SetHeight(3000000)
	table.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000)

	headers := []string{"Feature", "Status", "Notes"}
	for i, h := range headers {
		table.GetCell(0, i).SetText(h)
		table.GetCell(0, i).SetFill(gp.NewFill().SetSolid(gp.NewColor("4472C4")))
	}
	table.GetCell(1, 0).SetText("Rich Text")
	table.GetCell(1, 1).SetText("Done")
	table.GetCell(1, 2).SetText("Full support")
	table.GetCell(2, 0).SetText("Images")
	table.GetCell(2, 1).SetText("Done")
	table.GetCell(2, 2).SetText("PNG, JPEG, GIF")

	// Save
	w, _ := gp.NewWriter(pres, gp.WriterPowerPoint2007)
	err := w.Save("example_output.pptx")
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		return
	}
	defer os.Remove("example_output.pptx")

	fmt.Println("Presentation created successfully!")
	fmt.Printf("Slides: %d\n", pres.GetSlideCount())

	// Read it back
	reader, _ := gp.NewReader(gp.ReaderPowerPoint2007)
	readPres, err := reader.Read("example_output.pptx")
	if err != nil {
		fmt.Printf("Read error: %v\n", err)
		return
	}
	fmt.Printf("Read back: %d slides\n", readPres.GetSlideCount())

	// Output:
	// Presentation created successfully!
	// Slides: 3
	// Read back: 3 slides
}
