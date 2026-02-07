package gopresentation

import "testing"

func TestSlideBasics(t *testing.T) {
	s := newSlide()

	// Name
	s.SetName("Test Slide")
	if s.GetName() != "Test Slide" {
		t.Errorf("expected name 'Test Slide', got '%s'", s.GetName())
	}

	// Notes
	s.SetNotes("Speaker notes here")
	if s.GetNotes() != "Speaker notes here" {
		t.Errorf("expected notes, got '%s'", s.GetNotes())
	}

	// Visibility
	if !s.IsVisible() {
		t.Error("slide should be visible by default")
	}
	s.SetVisible(false)
	if s.IsVisible() {
		t.Error("slide should be hidden")
	}

	// Transition
	if s.GetTransition() != nil {
		t.Error("transition should be nil by default")
	}
	trans := &Transition{
		Type:     TransitionFade,
		Speed:    TransitionSpeedMedium,
		Duration: 500,
	}
	s.SetTransition(trans)
	if s.GetTransition().Type != TransitionFade {
		t.Error("transition type should be Fade")
	}
}

func TestSlideShapeCreation(t *testing.T) {
	s := newSlide()

	// Rich text
	rt := s.CreateRichTextShape()
	if rt == nil {
		t.Fatal("CreateRichTextShape() returned nil")
	}
	if rt.GetType() != ShapeTypeRichText {
		t.Error("expected ShapeTypeRichText")
	}

	// Drawing
	ds := s.CreateDrawingShape()
	if ds == nil {
		t.Fatal("CreateDrawingShape() returned nil")
	}
	if ds.GetType() != ShapeTypeDrawing {
		t.Error("expected ShapeTypeDrawing")
	}

	// Table
	ts := s.CreateTableShape(3, 4)
	if ts == nil {
		t.Fatal("CreateTableShape() returned nil")
	}
	if ts.GetType() != ShapeTypeTable {
		t.Error("expected ShapeTypeTable")
	}

	// Auto shape
	as := s.CreateAutoShape()
	if as == nil {
		t.Fatal("CreateAutoShape() returned nil")
	}
	if as.GetType() != ShapeTypeAutoShape {
		t.Error("expected ShapeTypeAutoShape")
	}

	// Line
	ls := s.CreateLineShape()
	if ls == nil {
		t.Fatal("CreateLineShape() returned nil")
	}
	if ls.GetType() != ShapeTypeLine {
		t.Error("expected ShapeTypeLine")
	}

	// Check shapes count
	shapes := s.GetShapes()
	if len(shapes) != 5 {
		t.Errorf("expected 5 shapes, got %d", len(shapes))
	}

	// Remove shape
	err := s.RemoveShape(0)
	if err != nil {
		t.Fatalf("RemoveShape(0) error: %v", err)
	}
	if len(s.GetShapes()) != 4 {
		t.Errorf("expected 4 shapes after removal, got %d", len(s.GetShapes()))
	}

	// Remove out of range
	err = s.RemoveShape(100)
	if err == nil {
		t.Error("expected error for out of range removal")
	}
}

func TestSlideAddShape(t *testing.T) {
	s := newSlide()
	shape := NewAutoShape()
	s.AddShape(shape)
	if len(s.GetShapes()) != 1 {
		t.Errorf("expected 1 shape, got %d", len(s.GetShapes()))
	}
}
