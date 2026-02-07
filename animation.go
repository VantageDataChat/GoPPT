package gopresentation

// Animation represents a slide animation.
type Animation struct {
	ShapeIndexes []int // indexes of shapes in this animation group
}

// NewAnimation creates a new animation.
func NewAnimation() *Animation {
	return &Animation{
		ShapeIndexes: make([]int, 0),
	}
}

// AddShape adds a shape index to the animation.
func (a *Animation) AddShape(index int) *Animation {
	a.ShapeIndexes = append(a.ShapeIndexes, index)
	return a
}

// GetShapeIndexes returns the shape indexes.
func (a *Animation) GetShapeIndexes() []int {
	return a.ShapeIndexes
}
