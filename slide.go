package gopresentation

import "errors"

// Transition represents a slide transition.
type Transition struct {
	Type     TransitionType
	Speed    TransitionSpeed
	Duration int // in milliseconds
}

// TransitionType represents the type of slide transition.
type TransitionType int

const (
	TransitionNone TransitionType = iota
	TransitionFade
	TransitionPush
	TransitionWipe
	TransitionSplit
	TransitionCover
	TransitionUncover
	TransitionDissolve
)

// TransitionSpeed represents the speed of a transition.
type TransitionSpeed string

const (
	TransitionSpeedSlow   TransitionSpeed = "slow"
	TransitionSpeedMedium TransitionSpeed = "med"
	TransitionSpeedFast   TransitionSpeed = "fast"
)

// Slide represents a single slide in a presentation.
type Slide struct {
	shapes     []Shape
	name       string
	notes      string
	transition *Transition
	visible    bool
	comments   []*Comment
	animations []*Animation
	background *Fill
}

// newSlide creates a new empty slide.
func newSlide() *Slide {
	return &Slide{
		shapes:     make([]Shape, 0),
		visible:    true,
		comments:   make([]*Comment, 0),
		animations: make([]*Animation, 0),
	}
}

// GetName returns the slide name.
func (s *Slide) GetName() string {
	return s.name
}

// SetName sets the slide name.
func (s *Slide) SetName(name string) {
	s.name = name
}

// GetNotes returns the slide notes.
func (s *Slide) GetNotes() string {
	return s.notes
}

// SetNotes sets the slide notes.
func (s *Slide) SetNotes(notes string) {
	s.notes = notes
}

// IsVisible returns whether the slide is visible.
func (s *Slide) IsVisible() bool {
	return s.visible
}

// SetVisible sets the slide visibility.
func (s *Slide) SetVisible(visible bool) {
	s.visible = visible
}

// GetTransition returns the slide transition.
func (s *Slide) GetTransition() *Transition {
	return s.transition
}

// SetTransition sets the slide transition.
func (s *Slide) SetTransition(t *Transition) {
	s.transition = t
}

// GetShapes returns all shapes on the slide.
func (s *Slide) GetShapes() []Shape {
	return s.shapes
}

// AddShape adds a shape to the slide.
func (s *Slide) AddShape(shape Shape) {
	s.shapes = append(s.shapes, shape)
}

// RemoveShape removes a shape by index.
func (s *Slide) RemoveShape(index int) error {
	if index < 0 || index >= len(s.shapes) {
		return errors.New("shape index out of range")
	}
	s.shapes = append(s.shapes[:index], s.shapes[index+1:]...)
	return nil
}

// CreateRichTextShape creates a new rich text shape and adds it to the slide.
func (s *Slide) CreateRichTextShape() *RichTextShape {
	shape := NewRichTextShape()
	s.shapes = append(s.shapes, shape)
	return shape
}

// CreateDrawingShape creates a new drawing (image) shape and adds it to the slide.
func (s *Slide) CreateDrawingShape() *DrawingShape {
	shape := NewDrawingShape()
	s.shapes = append(s.shapes, shape)
	return shape
}

// CreateTableShape creates a new table shape and adds it to the slide.
func (s *Slide) CreateTableShape(rows, cols int) *TableShape {
	shape := NewTableShape(rows, cols)
	s.shapes = append(s.shapes, shape)
	return shape
}

// CreateAutoShape creates a new auto shape and adds it to the slide.
func (s *Slide) CreateAutoShape() *AutoShape {
	shape := NewAutoShape()
	s.shapes = append(s.shapes, shape)
	return shape
}

// CreateLineShape creates a new line shape and adds it to the slide.
func (s *Slide) CreateLineShape() *LineShape {
	shape := NewLineShape()
	s.shapes = append(s.shapes, shape)
	return shape
}

// CreateChartShape creates a new chart shape and adds it to the slide.
func (s *Slide) CreateChartShape() *ChartShape {
	shape := NewChartShape()
	s.shapes = append(s.shapes, shape)
	return shape
}

// CreateGroupShape creates a new group shape and adds it to the slide.
func (s *Slide) CreateGroupShape() *GroupShape {
	shape := NewGroupShape()
	s.shapes = append(s.shapes, shape)
	return shape
}

// CreatePlaceholderShape creates a new placeholder shape and adds it to the slide.
func (s *Slide) CreatePlaceholderShape(phType PlaceholderType) *PlaceholderShape {
	shape := NewPlaceholderShape(phType)
	s.shapes = append(s.shapes, shape)
	return shape
}

// --- Comments ---

// AddComment adds a comment to the slide.
func (s *Slide) AddComment(c *Comment) {
	s.comments = append(s.comments, c)
}

// GetComments returns all comments on the slide.
func (s *Slide) GetComments() []*Comment {
	return s.comments
}

// GetCommentCount returns the number of comments.
func (s *Slide) GetCommentCount() int {
	return len(s.comments)
}

// --- Animations ---

// AddAnimation adds an animation to the slide.
func (s *Slide) AddAnimation(a *Animation) {
	s.animations = append(s.animations, a)
}

// GetAnimations returns all animations on the slide.
func (s *Slide) GetAnimations() []*Animation {
	return s.animations
}

// --- Background ---

// SetBackground sets the slide background fill.
func (s *Slide) SetBackground(f *Fill) {
	s.background = f
}

// GetBackground returns the slide background fill.
func (s *Slide) GetBackground() *Fill {
	return s.background
}
