package gopresentation

import (
	"fmt"
	"strings"
)

// Validate checks the presentation for structural issues and returns an error
// describing all problems found, or nil if the presentation is valid.
// This is analogous to unioffice's Validate() method.
func (p *Presentation) Validate() error {
	var errs []string

	if p.properties == nil {
		errs = append(errs, "document properties are nil")
	}
	if p.presentationProperties == nil {
		errs = append(errs, "presentation properties are nil")
	}
	if p.layout == nil {
		errs = append(errs, "document layout is nil")
	} else {
		if p.layout.CX <= 0 {
			errs = append(errs, "layout width (CX) must be positive")
		}
		if p.layout.CY <= 0 {
			errs = append(errs, "layout height (CY) must be positive")
		}
	}
	if len(p.slides) == 0 {
		errs = append(errs, "presentation must have at least one slide")
	}

	for i, slide := range p.slides {
		prefix := fmt.Sprintf("slide %d", i+1)
		if slideErrs := validateSlide(slide); len(slideErrs) > 0 {
			for _, e := range slideErrs {
				errs = append(errs, prefix+": "+e)
			}
		}
	}

	if len(errs) == 0 {
		return nil
	}
	return fmt.Errorf("validation failed:\n  %s", strings.Join(errs, "\n  "))
}

func validateSlide(s *Slide) []string {
	var errs []string
	for j, shape := range s.shapes {
		prefix := fmt.Sprintf("shape %d", j+1)
		if shape == nil {
			errs = append(errs, prefix+": shape is nil")
			continue
		}
		if shape.GetWidth() < 0 {
			errs = append(errs, prefix+": width is negative")
		}
		if shape.GetHeight() < 0 {
			errs = append(errs, prefix+": height is negative")
		}

		switch sh := shape.(type) {
		case *DrawingShape:
			if sh.data == nil && sh.path == "" {
				errs = append(errs, prefix+": drawing shape has no image data or path")
			}
		case *TableShape:
			if sh.numRows <= 0 || sh.numCols <= 0 {
				errs = append(errs, prefix+": table must have at least 1 row and 1 column")
			}
		case *ChartShape:
			if sh.plotArea.chartType == nil {
				errs = append(errs, prefix+": chart shape has no chart type set")
			}
		}
	}

	for j, c := range s.comments {
		if c.Author == nil {
			errs = append(errs, fmt.Sprintf("comment %d: missing author", j+1))
		}
	}

	return errs
}
