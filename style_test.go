package gopresentation

import "testing"

func TestColor(t *testing.T) {
	// From 8-char ARGB
	c := NewColor("FFE06B20")
	if c.ARGB != "FFE06B20" {
		t.Errorf("expected 'FFE06B20', got '%s'", c.ARGB)
	}

	// From 6-char RGB (auto-prefix FF)
	c2 := NewColor("E06B20")
	if c2.ARGB != "FFE06B20" {
		t.Errorf("expected 'FFE06B20', got '%s'", c2.ARGB)
	}

	// With hash prefix
	c3 := NewColor("#FF0000")
	if c3.ARGB != "FFFF0000" {
		t.Errorf("expected 'FFFF0000', got '%s'", c3.ARGB)
	}

	// Components
	c4 := NewColor("80FF8040")
	if c4.GetAlpha() != 0x80 {
		t.Errorf("expected alpha 128, got %d", c4.GetAlpha())
	}
	if c4.GetRed() != 0xFF {
		t.Errorf("expected red 255, got %d", c4.GetRed())
	}
	if c4.GetGreen() != 0x80 {
		t.Errorf("expected green 128, got %d", c4.GetGreen())
	}
	if c4.GetBlue() != 0x40 {
		t.Errorf("expected blue 64, got %d", c4.GetBlue())
	}

	// Predefined colors
	if ColorBlack.ARGB != "FF000000" {
		t.Error("ColorBlack mismatch")
	}
	if ColorWhite.ARGB != "FFFFFFFF" {
		t.Error("ColorWhite mismatch")
	}
}

func TestFont(t *testing.T) {
	f := NewFont()

	if f.Name != "Calibri" {
		t.Errorf("expected 'Calibri', got '%s'", f.Name)
	}
	if f.Size != 10 {
		t.Errorf("expected size 10, got %d", f.Size)
	}
	if f.Bold {
		t.Error("should not be bold by default")
	}

	// Chaining
	f.SetBold(true).SetItalic(true).SetSize(24).SetColor(ColorRed).SetName("Arial")
	if !f.Bold {
		t.Error("should be bold")
	}
	if !f.Italic {
		t.Error("should be italic")
	}
	if f.Size != 24 {
		t.Errorf("expected size 24, got %d", f.Size)
	}
	if f.Color != ColorRed {
		t.Error("color should be red")
	}
	if f.Name != "Arial" {
		t.Errorf("expected 'Arial', got '%s'", f.Name)
	}

	f.SetUnderline(UnderlineSingle)
	if f.Underline != UnderlineSingle {
		t.Error("expected single underline")
	}

	f.SetStrikethrough(true)
	if !f.Strikethrough {
		t.Error("should be strikethrough")
	}
}

func TestAlignment(t *testing.T) {
	a := NewAlignment()

	if a.Horizontal != HorizontalLeft {
		t.Error("expected left alignment")
	}
	if a.Vertical != VerticalTop {
		t.Error("expected top alignment")
	}

	a.SetHorizontal(HorizontalCenter).SetVertical(VerticalMiddle)
	if a.Horizontal != HorizontalCenter {
		t.Error("expected center alignment")
	}
	if a.Vertical != VerticalMiddle {
		t.Error("expected middle alignment")
	}
}

func TestFill(t *testing.T) {
	f := NewFill()
	if f.Type != FillNone {
		t.Error("expected FillNone")
	}

	// Solid
	f.SetSolid(ColorBlue)
	if f.Type != FillSolid {
		t.Error("expected FillSolid")
	}
	if f.Color != ColorBlue {
		t.Error("expected blue color")
	}

	// Gradient
	f2 := NewFill()
	f2.SetGradientLinear(ColorRed, ColorBlue, 90)
	if f2.Type != FillGradientLinear {
		t.Error("expected FillGradientLinear")
	}
	if f2.Rotation != 90 {
		t.Errorf("expected rotation 90, got %d", f2.Rotation)
	}
}

func TestBorder(t *testing.T) {
	b := NewBorder()
	if b.Style != BorderNone {
		t.Error("expected BorderNone")
	}
}

func TestShadow(t *testing.T) {
	s := NewShadow()
	if s.Visible {
		t.Error("shadow should not be visible by default")
	}

	s.SetVisible(true).SetDirection(45).SetDistance(10)
	if !s.Visible {
		t.Error("shadow should be visible")
	}
	if s.Direction != 45 {
		t.Errorf("expected direction 45, got %d", s.Direction)
	}
	if s.Distance != 10 {
		t.Errorf("expected distance 10, got %d", s.Distance)
	}
}

func TestHyperlink(t *testing.T) {
	// External
	hl := NewHyperlink("https://example.com")
	if hl.URL != "https://example.com" {
		t.Error("URL mismatch")
	}
	if hl.IsInternal {
		t.Error("should not be internal")
	}

	// Internal
	ihl := NewInternalHyperlink(3)
	if !ihl.IsInternal {
		t.Error("should be internal")
	}
	if ihl.SlideNumber != 3 {
		t.Errorf("expected slide 3, got %d", ihl.SlideNumber)
	}
}
