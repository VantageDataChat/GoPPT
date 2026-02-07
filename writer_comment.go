package gopresentation

import (
	"archive/zip"
	"fmt"
)

func (w *PPTXWriter) hasComments() bool {
	for _, slide := range w.presentation.slides {
		if len(slide.comments) > 0 {
			return true
		}
	}
	return false
}

func (w *PPTXWriter) collectAuthors() []*CommentAuthor {
	seen := make(map[string]*CommentAuthor)
	var authors []*CommentAuthor
	id := 0

	for _, slide := range w.presentation.slides {
		for _, c := range slide.comments {
			if c.Author != nil {
				if _, ok := seen[c.Author.Name]; !ok {
					c.Author.ID = id
					c.Author.ColorIdx = id
					seen[c.Author.Name] = c.Author
					authors = append(authors, c.Author)
					id++
				} else {
					c.Author.ID = seen[c.Author.Name].ID
				}
			}
		}
	}
	return authors
}

func (w *PPTXWriter) writeCommentAuthors(zw *zip.Writer) error {
	if !w.hasComments() {
		return nil
	}

	authors := w.collectAuthors()
	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmAuthorLst xmlns:p="%s">`, nsPresentationML)

	for _, a := range authors {
		content += fmt.Sprintf(`
  <p:cmAuthor id="%d" name="%s" initials="%s" lastIdx="0" clrIdx="%d"/>`,
			a.ID, xmlEscape(a.Name), xmlEscape(a.Initials), a.ColorIdx)
	}
	content += `
</p:cmAuthorLst>`

	return writeRawXMLToZip(zw, "ppt/commentAuthors.xml", content)
}

func (w *PPTXWriter) writeSlideComments(zw *zip.Writer, slide *Slide, slideNum int) error {
	if len(slide.comments) == 0 {
		return nil
	}

	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmLst xmlns:p="%s">`, nsPresentationML)

	for idx, c := range slide.comments {
		authorID := 0
		if c.Author != nil {
			authorID = c.Author.ID
		}
		content += fmt.Sprintf(`
  <p:cm authorId="%d" dt="%s" idx="%d">
    <p:pos x="%d" y="%d"/>
    <p:text>%s</p:text>
  </p:cm>`,
			authorID,
			c.Date.UTC().Format("2006-01-02T15:04:05.000"),
			idx+1,
			c.PositionX, c.PositionY,
			xmlEscape(c.Text))
	}
	content += `
</p:cmLst>`

	return writeRawXMLToZip(zw, fmt.Sprintf("ppt/comments/comment%d.xml", slideNum), content)
}
