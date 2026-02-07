package gopresentation

import "time"

// Comment represents a slide comment.
type Comment struct {
	Author    *CommentAuthor
	Text      string
	Date      time.Time
	PositionX int // in 1/100th of a point
	PositionY int
}

// CommentAuthor represents a comment author.
type CommentAuthor struct {
	Name     string
	Initials string
	ID       int
	ColorIdx int
}

// NewComment creates a new comment.
func NewComment() *Comment {
	return &Comment{
		Date: time.Now(),
	}
}

// SetAuthor sets the comment author.
func (c *Comment) SetAuthor(a *CommentAuthor) *Comment {
	c.Author = a
	return c
}

// SetText sets the comment text.
func (c *Comment) SetText(text string) *Comment {
	c.Text = text
	return c
}

// SetPosition sets the comment position.
func (c *Comment) SetPosition(x, y int) *Comment {
	c.PositionX = x
	c.PositionY = y
	return c
}

// SetDate sets the comment date.
func (c *Comment) SetDate(d time.Time) *Comment {
	c.Date = d
	return c
}

// NewCommentAuthor creates a new comment author.
func NewCommentAuthor(name, initials string) *CommentAuthor {
	return &CommentAuthor{
		Name:     name,
		Initials: initials,
	}
}
