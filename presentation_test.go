package gopresentation

import (
	"testing"
	"time"
)

func TestNewPresentation(t *testing.T) {
	p := New()
	if p == nil {
		t.Fatal("New() returned nil")
	}
	if p.GetSlideCount() != 1 {
		t.Errorf("expected 1 slide, got %d", p.GetSlideCount())
	}
	if p.GetActiveSlideIndex() != 0 {
		t.Errorf("expected active slide index 0, got %d", p.GetActiveSlideIndex())
	}
	if p.GetActiveSlide() == nil {
		t.Error("GetActiveSlide() returned nil")
	}
}

func TestPresentationSlideManagement(t *testing.T) {
	p := New()

	// Create additional slides
	s2 := p.CreateSlide()
	if s2 == nil {
		t.Fatal("CreateSlide() returned nil")
	}
	if p.GetSlideCount() != 2 {
		t.Errorf("expected 2 slides, got %d", p.GetSlideCount())
	}

	s3 := p.CreateSlide()
	if p.GetSlideCount() != 3 {
		t.Errorf("expected 3 slides, got %d", p.GetSlideCount())
	}

	// Get slide by index
	slide, err := p.GetSlide(1)
	if err != nil {
		t.Fatalf("GetSlide(1) error: %v", err)
	}
	if slide != s2 {
		t.Error("GetSlide(1) returned wrong slide")
	}

	// Out of range
	_, err = p.GetSlide(10)
	if err == nil {
		t.Error("expected error for out of range index")
	}

	// Set active slide
	err = p.SetActiveSlideIndex(2)
	if err != nil {
		t.Fatalf("SetActiveSlideIndex(2) error: %v", err)
	}
	if p.GetActiveSlide() != s3 {
		t.Error("active slide should be s3")
	}

	// Remove slide
	err = p.RemoveSlideByIndex(1)
	if err != nil {
		t.Fatalf("RemoveSlideByIndex(1) error: %v", err)
	}
	if p.GetSlideCount() != 2 {
		t.Errorf("expected 2 slides after removal, got %d", p.GetSlideCount())
	}

	// Remove out of range
	err = p.RemoveSlideByIndex(10)
	if err == nil {
		t.Error("expected error for out of range removal")
	}

	// Set active slide out of range
	err = p.SetActiveSlideIndex(10)
	if err == nil {
		t.Error("expected error for out of range active slide")
	}

	// GetAllSlides
	all := p.GetAllSlides()
	if len(all) != 2 {
		t.Errorf("expected 2 slides, got %d", len(all))
	}
}

func TestPresentationAddSlide(t *testing.T) {
	p := New()
	slide := newSlide()
	slide.SetName("External Slide")
	p.AddSlide(slide)
	if p.GetSlideCount() != 2 {
		t.Errorf("expected 2 slides, got %d", p.GetSlideCount())
	}
}

func TestDocumentProperties(t *testing.T) {
	p := New()
	props := p.GetDocumentProperties()

	props.Creator = "Test Author"
	props.Title = "Test Title"
	props.Description = "Test Description"
	props.Subject = "Test Subject"
	props.Keywords = "test, keywords"
	props.Category = "Test Category"
	props.Company = "Test Company"
	props.LastModifiedBy = "Test Modifier"
	props.Status = "Draft"
	props.Revision = "1.0"

	now := time.Now()
	props.Created = now
	props.Modified = now

	if props.Creator != "Test Author" {
		t.Errorf("expected creator 'Test Author', got '%s'", props.Creator)
	}
	if props.Title != "Test Title" {
		t.Errorf("expected title 'Test Title', got '%s'", props.Title)
	}
	if props.Description != "Test Description" {
		t.Errorf("expected description 'Test Description', got '%s'", props.Description)
	}
	if props.Company != "Test Company" {
		t.Errorf("expected company 'Test Company', got '%s'", props.Company)
	}
}

func TestCustomProperties(t *testing.T) {
	props := NewDocumentProperties()

	// Set custom properties
	props.SetCustomProperty("stringProp", "hello", PropertyTypeString)
	props.SetCustomProperty("intProp", 42, PropertyTypeInteger)
	props.SetCustomProperty("boolProp", true, PropertyTypeBoolean)
	props.SetCustomProperty("floatProp", 3.14, PropertyTypeFloat)

	// Check existence
	if !props.IsCustomPropertySet("stringProp") {
		t.Error("stringProp should be set")
	}
	if props.IsCustomPropertySet("nonexistent") {
		t.Error("nonexistent should not be set")
	}

	// Get values
	if v := props.GetCustomPropertyValue("stringProp"); v != "hello" {
		t.Errorf("expected 'hello', got '%v'", v)
	}
	if v := props.GetCustomPropertyValue("intProp"); v != 42 {
		t.Errorf("expected 42, got '%v'", v)
	}
	if v := props.GetCustomPropertyValue("nonexistent"); v != nil {
		t.Errorf("expected nil, got '%v'", v)
	}

	// Get types
	if pt := props.GetCustomPropertyType("stringProp"); pt != PropertyTypeString {
		t.Errorf("expected PropertyTypeString, got %v", pt)
	}
	if pt := props.GetCustomPropertyType("nonexistent"); pt != PropertyTypeUnknown {
		t.Errorf("expected PropertyTypeUnknown, got %v", pt)
	}

	// Get all property names
	names := props.GetCustomProperties()
	if len(names) != 4 {
		t.Errorf("expected 4 custom properties, got %d", len(names))
	}
}

func TestPresentationProperties(t *testing.T) {
	pp := NewPresentationProperties()

	// Zoom
	if pp.GetZoom() != 1.0 {
		t.Errorf("expected zoom 1.0, got %f", pp.GetZoom())
	}
	pp.SetZoom(2.5)
	if pp.GetZoom() != 2.5 {
		t.Errorf("expected zoom 2.5, got %f", pp.GetZoom())
	}

	// Last view
	if pp.GetLastView() != ViewSlide {
		t.Errorf("expected ViewSlide, got %v", pp.GetLastView())
	}
	pp.SetLastView(ViewNotes)
	if pp.GetLastView() != ViewNotes {
		t.Errorf("expected ViewNotes, got %v", pp.GetLastView())
	}

	// Slideshow type
	if pp.GetSlideshowType() != SlideshowTypePresent {
		t.Errorf("expected SlideshowTypePresent, got %v", pp.GetSlideshowType())
	}
	pp.SetSlideshowType(SlideshowTypeKiosk)
	if pp.GetSlideshowType() != SlideshowTypeKiosk {
		t.Errorf("expected SlideshowTypeKiosk, got %v", pp.GetSlideshowType())
	}

	// Comment visibility
	if pp.IsCommentVisible() {
		t.Error("comments should not be visible by default")
	}
	pp.SetCommentVisible(true)
	if !pp.IsCommentVisible() {
		t.Error("comments should be visible after setting")
	}

	// Mark as final
	if pp.IsMarkedAsFinal() {
		t.Error("should not be marked as final by default")
	}
	pp.MarkAsFinal()
	if !pp.IsMarkedAsFinal() {
		t.Error("should be marked as final")
	}
	pp.MarkAsFinal(false)
	if pp.IsMarkedAsFinal() {
		t.Error("should not be marked as final after MarkAsFinal(false)")
	}

	// Thumbnail
	pp.SetThumbnailPath("/path/to/thumb.png")
	if pp.GetThumbnailPath() != "/path/to/thumb.png" {
		t.Errorf("expected thumbnail path, got '%s'", pp.GetThumbnailPath())
	}
	pp.SetThumbnailData([]byte{1, 2, 3})
	if len(pp.GetThumbnailData()) != 3 {
		t.Errorf("expected 3 bytes, got %d", len(pp.GetThumbnailData()))
	}
}

func TestDocumentLayout(t *testing.T) {
	dl := NewDocumentLayout()

	// Default 4:3
	if dl.CX != 9144000 {
		t.Errorf("expected CX 9144000, got %d", dl.CX)
	}
	if dl.CY != 6858000 {
		t.Errorf("expected CY 6858000, got %d", dl.CY)
	}
	if dl.Name != LayoutScreen4x3 {
		t.Errorf("expected layout name '%s', got '%s'", LayoutScreen4x3, dl.Name)
	}

	// 16:9
	dl.SetLayout(LayoutScreen16x9)
	if dl.CX != 12192000 {
		t.Errorf("expected CX 12192000, got %d", dl.CX)
	}

	// Custom
	dl.SetCustomLayout(5000000, 3000000)
	if dl.CX != 5000000 || dl.CY != 3000000 {
		t.Errorf("expected custom layout 5000000x3000000, got %dx%d", dl.CX, dl.CY)
	}
	if dl.Name != LayoutCustom {
		t.Errorf("expected layout name '%s', got '%s'", LayoutCustom, dl.Name)
	}
}

func TestSlideMaster(t *testing.T) {
	p := New()
	sm := p.CreateSlideMaster()
	if sm == nil {
		t.Fatal("CreateSlideMaster() returned nil")
	}
	masters := p.GetSlideMasters()
	if len(masters) != 1 {
		t.Errorf("expected 1 slide master, got %d", len(masters))
	}
}

func TestSetDocumentProperties(t *testing.T) {
	p := New()
	newProps := NewDocumentProperties()
	newProps.Creator = "New Creator"
	p.SetDocumentProperties(newProps)
	if p.GetDocumentProperties().Creator != "New Creator" {
		t.Error("SetDocumentProperties did not work")
	}
}

func TestSetLayout(t *testing.T) {
	p := New()
	layout := NewDocumentLayout()
	layout.SetLayout(LayoutScreen16x9)
	p.SetLayout(layout)
	if p.GetLayout().CX != 12192000 {
		t.Error("SetLayout did not work")
	}
}

func TestRemoveSlideAdjustsActiveIndex(t *testing.T) {
	p := New()
	p.CreateSlide()
	p.CreateSlide()
	p.SetActiveSlideIndex(2)
	p.RemoveSlideByIndex(2)
	if p.GetActiveSlideIndex() >= p.GetSlideCount() {
		t.Error("active slide index should be adjusted after removal")
	}
}
