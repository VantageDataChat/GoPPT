package gopresentation

import (
	"bytes"
	"fmt"
	"image"
	"image/color"
	"image/draw"
	_ "image/gif"
	"image/jpeg"
	"image/png"
	"math"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"golang.org/x/image/font"
	"golang.org/x/image/font/basicfont"
	"golang.org/x/image/math/fixed"
)

// ImageFormat represents the output image format.
type ImageFormat int

const (
	ImageFormatPNG ImageFormat = iota
	ImageFormatJPEG
)

// RenderOptions configures slide-to-image rendering.
type RenderOptions struct {
	// Width is the output image width in pixels. Height is calculated from slide aspect ratio.
	// Default: 960
	Width int
	// Format is the output image format (PNG or JPEG).
	Format ImageFormat
	// JPEGQuality is the JPEG quality (1-100). Default: 90.
	JPEGQuality int
	// BackgroundColor overrides the slide background. Nil means use slide background or white.
	BackgroundColor *color.RGBA
	// DPI is the rendering DPI for font sizing. Default: 96.
	DPI float64
	// FontDirs specifies additional directories to search for TrueType/OpenType fonts.
	// System font directories are always searched automatically.
	FontDirs []string
	// FontCache allows sharing a pre-configured FontCache across multiple renders.
	// If nil, a new FontCache is created using FontDirs.
	FontCache *FontCache
}

// DefaultRenderOptions returns default rendering options.
func DefaultRenderOptions() *RenderOptions {
	return &RenderOptions{
		Width:       960,
		Format:      ImageFormatPNG,
		JPEGQuality: 90,
		DPI:         96,
	}
}

// SlideToImage renders a single slide to an image.
func (p *Presentation) SlideToImage(slideIndex int, opts *RenderOptions) (image.Image, error) {
	if slideIndex < 0 || slideIndex >= len(p.slides) {
		return nil, fmt.Errorf("slide index %d out of range (0-%d)", slideIndex, len(p.slides)-1)
	}
	if opts == nil {
		opts = DefaultRenderOptions()
	}
	if opts.Width <= 0 {
		opts.Width = 960
	}

	slide := p.slides[slideIndex]
	layout := p.layout

	slideW := float64(layout.CX)
	slideH := float64(layout.CY)
	imgW := opts.Width
	imgH := int(float64(imgW) * slideH / slideW)

	scaleX := float64(imgW) / slideW
	scaleY := float64(imgH) / slideH

	img := image.NewRGBA(image.Rect(0, 0, imgW, imgH))

	fc := opts.FontCache
	if fc == nil {
		fc = NewFontCache(opts.FontDirs...)
	}
	dpi := opts.DPI
	if dpi <= 0 {
		dpi = 96
	}

	r := &renderer{
		img:       img,
		scaleX:    scaleX,
		scaleY:    scaleY,
		fontCache: fc,
		dpi:       dpi,
	}

	// Fill background
	bgColor := color.RGBA{R: 255, G: 255, B: 255, A: 255}
	drawn := false
	if opts.BackgroundColor != nil {
		bgColor = *opts.BackgroundColor
	} else if slide.background != nil {
		switch slide.background.Type {
		case FillSolid:
			bgColor = argbToRGBA(slide.background.Color)
		case FillGradientLinear:
			r.fillGradientLinear(img.Bounds(), slide.background)
			drawn = true
		case FillGradientPath:
			r.fillGradientPath(img.Bounds(), slide.background)
			drawn = true
		}
	}
	if !drawn {
		r.fillRectFast(img.Bounds(), bgColor)
	}

	for _, shape := range slide.shapes {
		r.renderShape(shape)
	}

	return img, nil
}

// SlidesToImages renders all slides to images.
func (p *Presentation) SlidesToImages(opts *RenderOptions) ([]image.Image, error) {
	if opts == nil {
		opts = DefaultRenderOptions()
	}
	if opts.FontCache == nil {
		opts.FontCache = NewFontCache(opts.FontDirs...)
	}
	images := make([]image.Image, len(p.slides))
	for i := range p.slides {
		img, err := p.SlideToImage(i, opts)
		if err != nil {
			return nil, fmt.Errorf("slide %d: %w", i, err)
		}
		images[i] = img
	}
	return images, nil
}

// SaveSlideAsImage renders a slide and saves it to a file.
func (p *Presentation) SaveSlideAsImage(slideIndex int, path string, opts *RenderOptions) error {
	img, err := p.SlideToImage(slideIndex, opts)
	if err != nil {
		return err
	}
	return saveImage(img, path, opts)
}

// SaveSlidesAsImages renders all slides and saves them to files.
// The pattern should contain %d for the slide number (1-based), e.g. "slide_%d.png".
func (p *Presentation) SaveSlidesAsImages(pattern string, opts *RenderOptions) error {
	for i := range p.slides {
		path := fmt.Sprintf(pattern, i+1)
		if err := p.SaveSlideAsImage(i, path, opts); err != nil {
			return fmt.Errorf("slide %d: %w", i+1, err)
		}
	}
	return nil
}

func saveImage(img image.Image, path string, opts *RenderOptions) error {
	if opts == nil {
		opts = DefaultRenderOptions()
	}
	dir := filepath.Dir(path)
	if dir != "" && dir != "." {
		if err := os.MkdirAll(dir, 0750); err != nil {
			return fmt.Errorf("create directory: %w", err)
		}
	}
	f, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0644)
	if err != nil {
		return fmt.Errorf("create file: %w", err)
	}
	var encodeErr error
	switch opts.Format {
	case ImageFormatJPEG:
		quality := opts.JPEGQuality
		if quality <= 0 || quality > 100 {
			quality = 90
		}
		encodeErr = jpeg.Encode(f, img, &jpeg.Options{Quality: quality})
	default:
		encodeErr = png.Encode(f, img)
	}
	closeErr := f.Close()
	if encodeErr != nil {
		return encodeErr
	}
	return closeErr
}

// --- renderer core ---

type renderer struct {
	img       *image.RGBA
	scaleX    float64
	scaleY    float64
	fontCache *FontCache
	dpi       float64
}

func (r *renderer) renderShape(shape Shape) {
	switch s := shape.(type) {
	case *RichTextShape:
		r.renderRichText(s)
	case *PlaceholderShape:
		r.renderRichText(&s.RichTextShape)
	case *DrawingShape:
		r.renderDrawing(s)
	case *AutoShape:
		r.renderAutoShape(s)
	case *LineShape:
		r.renderLine(s)
	case *TableShape:
		r.renderTable(s)
	case *ChartShape:
		r.renderChart(s)
	case *GroupShape:
		r.renderGroup(s)
	}
}

func (r *renderer) emuToPixelX(emu int64) int { return int(float64(emu) * r.scaleX) }
func (r *renderer) emuToPixelY(emu int64) int { return int(float64(emu) * r.scaleY) }

func argbToRGBA(c Color) color.RGBA {
	return color.RGBA{R: c.GetRed(), G: c.GetGreen(), B: c.GetBlue(), A: c.GetAlpha()}
}

// --- Pixel operations (performance-critical) ---

// blendPixel alpha-blends color c over the existing pixel at (x, y).
// Uses direct Pix slice access for performance.
func (r *renderer) blendPixel(x, y int, c color.RGBA) {
	b := r.img.Bounds()
	if x < b.Min.X || x >= b.Max.X || y < b.Min.Y || y >= b.Max.Y {
		return
	}
	if c.A == 0 {
		return
	}
	off := (y-b.Min.Y)*r.img.Stride + (x-b.Min.X)*4
	pix := r.img.Pix
	if c.A == 255 {
		pix[off] = c.R
		pix[off+1] = c.G
		pix[off+2] = c.B
		pix[off+3] = 255
		return
	}
	a := uint32(c.A)
	ia := 255 - a
	pix[off] = uint8((uint32(c.R)*a + uint32(pix[off])*ia) / 255)
	pix[off+1] = uint8((uint32(c.G)*a + uint32(pix[off+1])*ia) / 255)
	pix[off+2] = uint8((uint32(c.B)*a + uint32(pix[off+2])*ia) / 255)
	pix[off+3] = uint8(uint32(pix[off+3]) + (255-uint32(pix[off+3]))*a/255)
}

// blendPixelF blends with fractional coverage (0.0–1.0) for anti-aliasing.
func (r *renderer) blendPixelF(x, y int, c color.RGBA, coverage float64) {
	if coverage <= 0 {
		return
	}
	if coverage >= 1.0 {
		r.blendPixel(x, y, c)
		return
	}
	r.blendPixel(x, y, color.RGBA{R: c.R, G: c.G, B: c.B, A: uint8(float64(c.A) * coverage)})
}

// fillRectFast fills a rectangle with an opaque color using draw.Draw.
func (r *renderer) fillRectFast(rect image.Rectangle, c color.RGBA) {
	draw.Draw(r.img, rect, &image.Uniform{c}, image.Point{}, draw.Over)
}

// fillRectBlend fills a rectangle with alpha blending, using row-based direct Pix access.
func (r *renderer) fillRectBlend(rect image.Rectangle, c color.RGBA) {
	b := r.img.Bounds()
	rect = rect.Intersect(b)
	if rect.Empty() {
		return
	}
	if c.A == 0 {
		return
	}
	if c.A == 255 {
		r.fillRectFast(rect, c)
		return
	}
	a := uint32(c.A)
	ia := 255 - a
	cr, cg, cb := uint32(c.R)*a, uint32(c.G)*a, uint32(c.B)*a
	pix := r.img.Pix
	stride := r.img.Stride
	minX := rect.Min.X - b.Min.X
	minY := rect.Min.Y - b.Min.Y
	w := rect.Dx()
	for dy := 0; dy < rect.Dy(); dy++ {
		off := (minY+dy)*stride + minX*4
		for dx := 0; dx < w; dx++ {
			pix[off] = uint8((cr + uint32(pix[off])*ia) / 255)
			pix[off+1] = uint8((cg + uint32(pix[off+1])*ia) / 255)
			pix[off+2] = uint8((cb + uint32(pix[off+2])*ia) / 255)
			pix[off+3] = uint8(uint32(pix[off+3]) + (255-uint32(pix[off+3]))*a/255)
			off += 4
		}
	}
}

// --- Rotation & flip support ---

func rotatedBounds(cx, cy float64, w, h int, angleDeg int) image.Rectangle {
	rad := float64(angleDeg) * math.Pi / 180.0
	cos := math.Abs(math.Cos(rad))
	sin := math.Abs(math.Sin(rad))
	fw, fh := float64(w), float64(h)
	newW := fw*cos + fh*sin
	newH := fw*sin + fh*cos
	return image.Rect(
		int(cx-newW/2), int(cy-newH/2),
		int(cx+newW/2)+1, int(cy+newH/2)+1,
	)
}

func (r *renderer) renderRotated(x, y, w, h, rotation int, flipH, flipV bool, drawFn func(tmp *renderer)) {
	if w <= 0 || h <= 0 {
		return
	}
	tmp := image.NewRGBA(image.Rect(0, 0, w, h))
	tmpR := &renderer{img: tmp, scaleX: r.scaleX, scaleY: r.scaleY, fontCache: r.fontCache, dpi: r.dpi}
	drawFn(tmpR)

	// Apply flips using direct Pix access
	if flipH || flipV {
		flipped := image.NewRGBA(tmp.Bounds())
		stride := tmp.Stride
		for py := 0; py < h; py++ {
			sy := py
			if flipV {
				sy = h - 1 - py
			}
			srcRow := sy * stride
			dstRow := py * flipped.Stride
			if flipH {
				for px := 0; px < w; px++ {
					sx := w - 1 - px
					sOff := srcRow + sx*4
					dOff := dstRow + px*4
					copy(flipped.Pix[dOff:dOff+4], tmp.Pix[sOff:sOff+4])
				}
			} else {
				copy(flipped.Pix[dstRow:dstRow+w*4], tmp.Pix[srcRow:srcRow+w*4])
			}
		}
		tmp = flipped
	}

	if rotation == 0 {
		draw.Draw(r.img, image.Rect(x, y, x+w, y+h), tmp, image.Point{}, draw.Over)
		return
	}

	rad := float64(rotation) * math.Pi / 180.0
	cosA := math.Cos(rad)
	sinA := math.Sin(rad)
	cx := float64(w) / 2
	cy := float64(h) / 2
	destCX := float64(x) + cx
	destCY := float64(y) + cy

	bounds := rotatedBounds(destCX, destCY, w, h, rotation)
	imgBounds := r.img.Bounds()
	// Clamp to image bounds
	minDY := maxInt(bounds.Min.Y, imgBounds.Min.Y)
	maxDY := minInt(bounds.Max.Y, imgBounds.Max.Y)
	minDX := maxInt(bounds.Min.X, imgBounds.Min.X)
	maxDX := minInt(bounds.Max.X, imgBounds.Max.X)

	for dy := minDY; dy < maxDY; dy++ {
		ry := float64(dy) - destCY
		for dx := minDX; dx < maxDX; dx++ {
			rx := float64(dx) - destCX
			sx := rx*cosA + ry*sinA + cx
			sy := -rx*sinA + ry*cosA + cy
			ix, iy := int(sx), int(sy)
			if ix >= 0 && ix < w && iy >= 0 && iy < h {
				sOff := iy*tmp.Stride + ix*4
				if tmp.Pix[sOff+3] > 0 {
					r.blendPixel(dx, dy, color.RGBA{
						R: tmp.Pix[sOff], G: tmp.Pix[sOff+1],
						B: tmp.Pix[sOff+2], A: tmp.Pix[sOff+3],
					})
				}
			}
		}
	}
}

func (r *renderer) renderGroup(g *GroupShape) {
	rotation := g.GetRotation()
	flipH := g.GetFlipHorizontal()
	flipV := g.GetFlipVertical()
	if rotation == 0 && !flipH && !flipV {
		for _, gs := range g.shapes {
			r.renderShape(gs)
		}
		return
	}
	x := r.emuToPixelX(g.offsetX)
	y := r.emuToPixelY(g.offsetY)
	w := r.emuToPixelX(g.width)
	h := r.emuToPixelY(g.height)
	r.renderRotated(x, y, w, h, rotation, flipH, flipV, func(tmp *renderer) {
		for _, gs := range g.shapes {
			tmp.renderShape(gs)
		}
	})
}

// --- Shape rendering ---

func (r *renderer) renderRichText(s *RichTextShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)
	rotation := s.GetRotation()
	flipH := s.GetFlipHorizontal()
	flipV := s.GetFlipVertical()

	drawContent := func(tr *renderer) {
		ox, oy := x, y
		if tr != r {
			ox, oy = 0, 0
		}
		rect := image.Rect(ox, oy, ox+w, oy+h)

		// Shadow BEFORE fill (so shadow appears behind)
		if s.shadow != nil && s.shadow.Visible {
			tr.renderShadow(s.shadow, rect)
		}
		tr.renderFill(s.fill, rect)
		if s.border != nil && s.border.Style != BorderNone {
			pw := maxInt(int(float64(maxInt(s.border.Width, 1))*tr.scaleX), 1)
			tr.drawRectBorder(rect, argbToRGBA(s.border.Color), pw, s.border.Style)
		}
		tr.drawParagraphs(s.paragraphs, ox, oy, w, h, s.textAnchor)
	}

	if rotation != 0 || flipH || flipV {
		r.renderRotated(x, y, w, h, rotation, flipH, flipV, drawContent)
	} else {
		drawContent(r)
	}
}

func (r *renderer) renderDrawing(s *DrawingShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)

	imgData := s.data
	if len(imgData) == 0 && s.path != "" {
		if data, err := os.ReadFile(s.path); err == nil {
			imgData = data
		}
	}
	if len(imgData) == 0 {
		return
	}

	srcImg, _, err := image.Decode(bytes.NewReader(imgData))
	if err != nil {
		r.drawRect(image.Rect(x, y, x+w, y+h), color.RGBA{R: 200, G: 200, B: 200, A: 255}, 1)
		return
	}

	rotation := s.GetRotation()
	flipH := s.GetFlipHorizontal()
	flipV := s.GetFlipVertical()

	drawImg := func(tr *renderer) {
		ox, oy := x, y
		if tr != r {
			ox, oy = 0, 0
		}
		scaledImg := scaleImageBilinear(srcImg, w, h)
		draw.Draw(tr.img, image.Rect(ox, oy, ox+w, oy+h), scaledImg, image.Point{}, draw.Over)
	}

	if rotation != 0 || flipH || flipV {
		r.renderRotated(x, y, w, h, rotation, flipH, flipV, drawImg)
	} else {
		drawImg(r)
	}
}

func (r *renderer) renderAutoShape(s *AutoShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)
	rotation := s.GetRotation()
	flipH := s.GetFlipHorizontal()
	flipV := s.GetFlipVertical()

	drawContent := func(tr *renderer) {
		ox, oy := x, y
		if tr != r {
			ox, oy = 0, 0
		}
		rect := image.Rect(ox, oy, ox+w, oy+h)
		if s.shadow != nil && s.shadow.Visible {
			tr.renderShadow(s.shadow, rect)
		}
		tr.renderAutoShapeFill(s, ox, oy, w, h)
		tr.renderAutoShapeBorder(s, ox, oy, w, h)
		if s.text != "" {
			tr.drawStringCentered(s.text, tr.getFace(NewFont()), color.RGBA{A: 255}, rect)
		}
	}

	if rotation != 0 || flipH || flipV {
		r.renderRotated(x, y, w, h, rotation, flipH, flipV, drawContent)
	} else {
		drawContent(r)
	}
}

func (r *renderer) renderAutoShapeFill(s *AutoShape, x, y, w, h int) {
	if s.fill == nil || s.fill.Type == FillNone {
		return
	}
	fc := argbToRGBA(s.fill.Color)
	rect := image.Rect(x, y, x+w, y+h)

	switch s.shapeType {
	case AutoShapeEllipse:
		if s.fill.Type == FillSolid {
			r.fillEllipseAA(x, y, w, h, fc)
		} else {
			r.fillGradientLinear(rect, s.fill)
		}
	case AutoShapeRoundedRect:
		radius := minInt(w, h) / 5
		if s.fill.Type == FillSolid {
			r.fillRoundedRect(x, y, w, h, radius, fc)
		} else {
			r.fillGradientLinear(rect, s.fill)
		}
	case AutoShapeTriangle:
		r.fillTriangle(x, y, w, h, fc)
	case AutoShapeDiamond:
		r.fillDiamond(x, y, w, h, fc)
	case AutoShapeHexagon:
		r.fillHexagon(x, y, w, h, fc)
	case AutoShapePentagon:
		r.fillPentagon(x, y, w, h, fc)
	case AutoShapeArrowRight:
		r.fillArrowRight(x, y, w, h, fc)
	case AutoShapeArrowLeft:
		r.fillArrowLeft(x, y, w, h, fc)
	case AutoShapeArrowUp:
		r.fillArrowUp(x, y, w, h, fc)
	case AutoShapeArrowDown:
		r.fillArrowDown(x, y, w, h, fc)
	case AutoShapeStar5:
		r.fillStar(x, y, w, h, 5, fc)
	case AutoShapeStar4:
		r.fillStar(x, y, w, h, 4, fc)
	case AutoShapeHeart:
		r.fillHeart(x, y, w, h, fc)
	case AutoShapePlus:
		r.fillPlus(x, y, w, h, fc)
	default:
		r.renderFill(s.fill, rect)
	}
}

func (r *renderer) renderAutoShapeBorder(s *AutoShape, x, y, w, h int) {
	if s.border == nil || s.border.Style == BorderNone {
		return
	}
	bc := argbToRGBA(s.border.Color)
	pw := maxInt(int(float64(maxInt(s.border.Width, 1))*r.scaleX), 1)

	switch s.shapeType {
	case AutoShapeEllipse:
		r.drawEllipseAA(x, y, w, h, bc, pw)
	case AutoShapeRoundedRect:
		r.drawRoundedRect(x, y, w, h, minInt(w, h)/5, bc, pw)
	case AutoShapeTriangle:
		r.drawTriangle(x, y, w, h, bc, pw)
	case AutoShapeDiamond:
		r.drawDiamond(x, y, w, h, bc, pw)
	default:
		r.drawRectBorder(image.Rect(x, y, x+w, y+h), bc, pw, s.border.Style)
	}
}

func (r *renderer) renderLine(s *LineShape) {
	x1 := r.emuToPixelX(s.offsetX)
	y1 := r.emuToPixelY(s.offsetY)
	x2 := r.emuToPixelX(s.offsetX + s.width)
	y2 := r.emuToPixelY(s.offsetY + s.height)
	pw := maxInt(int(float64(maxInt(s.lineWidth, 1))*r.scaleX), 1)
	r.drawLineAA(x1, y1, x2, y2, argbToRGBA(s.lineColor), pw)
}

func (r *renderer) renderTable(s *TableShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)
	if s.numRows == 0 || s.numCols == 0 {
		return
	}
	cellW := w / s.numCols
	cellH := h / s.numRows
	pad := 3

	for row := 0; row < s.numRows; row++ {
		for col := 0; col < s.numCols; col++ {
			cx := x + col*cellW
			cy := y + row*cellH
			cellRect := image.Rect(cx, cy, cx+cellW, cy+cellH)
			cell := s.rows[row][col]
			r.renderFill(cell.fill, cellRect)
			if cell.border != nil {
				r.renderCellBorders(cell.border, cellRect)
			} else {
				r.drawRect(cellRect, color.RGBA{A: 255}, 1)
			}
			r.drawParagraphs(cell.paragraphs, cx+pad, cy+pad, cellW-2*pad, cellH-2*pad, TextAnchorNone)
		}
	}
}

func (r *renderer) renderCellBorders(cb *CellBorders, rect image.Rectangle) {
	drawBorder := func(b *Border, x1, y1, x2, y2 int) {
		if b == nil || b.Style == BorderNone {
			return
		}
		pw := maxInt(int(float64(b.Width)*r.scaleX), 1)
		r.drawLineThick(x1, y1, x2, y2, argbToRGBA(b.Color), pw)
	}
	drawBorder(cb.Top, rect.Min.X, rect.Min.Y, rect.Max.X, rect.Min.Y)
	drawBorder(cb.Bottom, rect.Min.X, rect.Max.Y-1, rect.Max.X, rect.Max.Y-1)
	drawBorder(cb.Left, rect.Min.X, rect.Min.Y, rect.Min.X, rect.Max.Y)
	drawBorder(cb.Right, rect.Max.X-1, rect.Min.Y, rect.Max.X-1, rect.Max.Y)
}

// --- Fill rendering ---

func (r *renderer) renderFill(fill *Fill, rect image.Rectangle) {
	if fill == nil || fill.Type == FillNone {
		return
	}
	switch fill.Type {
	case FillSolid:
		fc := argbToRGBA(fill.Color)
		r.fillRectBlend(rect, fc)
	case FillGradientLinear:
		r.fillGradientLinear(rect, fill)
	case FillGradientPath:
		r.fillGradientPath(rect, fill)
	}
}

func (r *renderer) fillGradientLinear(rect image.Rectangle, fill *Fill) {
	startC := argbToRGBA(fill.Color)
	endC := argbToRGBA(fill.EndColor)
	w := rect.Dx()
	h := rect.Dy()
	if w <= 0 || h <= 0 {
		return
	}
	rad := float64(fill.Rotation) * math.Pi / 180.0
	cosA := math.Cos(rad)
	sinA := math.Sin(rad)
	cx := float64(w) / 2
	cy := float64(h) / 2
	maxProj := math.Abs(cx*cosA) + math.Abs(cy*sinA)
	if maxProj < 1 {
		maxProj = 1
	}
	invMaxProj := 1.0 / (2 * maxProj)

	// Pre-compute row-independent part
	pix := r.img.Pix
	bounds := r.img.Bounds()
	stride := r.img.Stride

	for py := rect.Min.Y; py < rect.Max.Y; py++ {
		if py < bounds.Min.Y || py >= bounds.Max.Y {
			continue
		}
		dyf := float64(py-rect.Min.Y) - cy
		rowBase := dyf*sinA + maxProj
		off := (py-bounds.Min.Y)*stride + (maxInt(rect.Min.X, bounds.Min.X)-bounds.Min.X)*4
		for px := maxInt(rect.Min.X, bounds.Min.X); px < minInt(rect.Max.X, bounds.Max.X); px++ {
			dxf := float64(px-rect.Min.X) - cx
			t := (dxf*cosA + rowBase) * invMaxProj
			if t < 0 {
				t = 0
			} else if t > 1 {
				t = 1
			}
			it := 1 - t
			pix[off] = uint8(float64(startC.R)*it + float64(endC.R)*t)
			pix[off+1] = uint8(float64(startC.G)*it + float64(endC.G)*t)
			pix[off+2] = uint8(float64(startC.B)*it + float64(endC.B)*t)
			pix[off+3] = uint8(float64(startC.A)*it + float64(endC.A)*t)
			off += 4
		}
	}
}

func (r *renderer) fillGradientPath(rect image.Rectangle, fill *Fill) {
	startC := argbToRGBA(fill.Color)
	endC := argbToRGBA(fill.EndColor)
	w := rect.Dx()
	h := rect.Dy()
	if w <= 0 || h <= 0 {
		return
	}
	cx := float64(w) / 2
	cy := float64(h) / 2
	maxDist := math.Sqrt(cx*cx + cy*cy)
	if maxDist < 1 {
		maxDist = 1
	}
	invMaxDist := 1.0 / maxDist

	pix := r.img.Pix
	bounds := r.img.Bounds()
	stride := r.img.Stride

	for py := rect.Min.Y; py < rect.Max.Y; py++ {
		if py < bounds.Min.Y || py >= bounds.Max.Y {
			continue
		}
		dyf := float64(py-rect.Min.Y) - cy
		dy2 := dyf * dyf
		off := (py-bounds.Min.Y)*stride + (maxInt(rect.Min.X, bounds.Min.X)-bounds.Min.X)*4
		for px := maxInt(rect.Min.X, bounds.Min.X); px < minInt(rect.Max.X, bounds.Max.X); px++ {
			dxf := float64(px-rect.Min.X) - cx
			t := math.Sqrt(dxf*dxf+dy2) * invMaxDist
			if t > 1 {
				t = 1
			}
			it := 1 - t
			pix[off] = uint8(float64(startC.R)*it + float64(endC.R)*t)
			pix[off+1] = uint8(float64(startC.G)*it + float64(endC.G)*t)
			pix[off+2] = uint8(float64(startC.B)*it + float64(endC.B)*t)
			pix[off+3] = uint8(float64(startC.A)*it + float64(endC.A)*t)
			off += 4
		}
	}
}

func lerpColor(a, b color.RGBA, t float64) color.RGBA {
	it := 1 - t
	return color.RGBA{
		R: uint8(float64(a.R)*it + float64(b.R)*t),
		G: uint8(float64(a.G)*it + float64(b.G)*t),
		B: uint8(float64(a.B)*it + float64(b.B)*t),
		A: uint8(float64(a.A)*it + float64(b.A)*t),
	}
}

// --- Shadow rendering ---

func (r *renderer) renderShadow(shadow *Shadow, rect image.Rectangle) {
	if shadow == nil || !shadow.Visible {
		return
	}
	rad := float64(shadow.Direction) * math.Pi / 180.0
	dist := float64(shadow.Distance) * r.scaleX
	dx := int(dist * math.Cos(rad))
	dy := int(dist * math.Sin(rad))
	shadowColor := argbToRGBA(shadow.Color)
	shadowColor.A = uint8(float64(shadow.Alpha) * 255 / 100)
	shadowRect := rect.Add(image.Pt(dx, dy))

	blur := shadow.BlurRadius
	if blur <= 0 {
		r.fillRectBlend(shadowRect, shadowColor)
		return
	}

	// Box-blur approximation: render shadow at full alpha, then apply a simple
	// multi-pass box expansion with decreasing alpha from outside in.
	// We draw from outermost ring inward so inner pixels get the strongest alpha.
	steps := minInt(blur, 10)
	for i := steps; i >= 0; i-- {
		t := float64(i) / float64(steps)
		alpha := uint8(float64(shadowColor.A) * (1 - t*t)) // quadratic falloff
		c := color.RGBA{R: shadowColor.R, G: shadowColor.G, B: shadowColor.B, A: alpha}
		expanded := shadowRect.Inset(-i)
		// Only draw the ring (not the interior) for outer layers
		if i > 0 {
			inner := shadowRect.Inset(-(i - 1))
			// Top strip
			r.fillRectBlend(image.Rect(expanded.Min.X, expanded.Min.Y, expanded.Max.X, inner.Min.Y), c)
			// Bottom strip
			r.fillRectBlend(image.Rect(expanded.Min.X, inner.Max.Y, expanded.Max.X, expanded.Max.Y), c)
			// Left strip
			r.fillRectBlend(image.Rect(expanded.Min.X, inner.Min.Y, inner.Min.X, inner.Max.Y), c)
			// Right strip
			r.fillRectBlend(image.Rect(inner.Max.X, inner.Min.Y, expanded.Max.X, inner.Max.Y), c)
		} else {
			r.fillRectBlend(expanded, c)
		}
	}
}

// --- Drawing primitives ---

func (r *renderer) drawRect(rect image.Rectangle, c color.RGBA, width int) {
	for i := 0; i < width; i++ {
		// Top and bottom horizontal lines
		r.fillRectBlend(image.Rect(rect.Min.X, rect.Min.Y+i, rect.Max.X, rect.Min.Y+i+1), c)
		r.fillRectBlend(image.Rect(rect.Min.X, rect.Max.Y-1-i, rect.Max.X, rect.Max.Y-i), c)
		// Left and right vertical lines
		for y := rect.Min.Y; y < rect.Max.Y; y++ {
			r.blendPixel(rect.Min.X+i, y, c)
			r.blendPixel(rect.Max.X-1-i, y, c)
		}
	}
}

func (r *renderer) drawRectBorder(rect image.Rectangle, c color.RGBA, width int, style BorderStyle) {
	if style == BorderSolid || style == BorderNone {
		r.drawRect(rect, c, width)
		return
	}
	dashLen, gapLen := 6, 4
	if style == BorderDot {
		dashLen, gapLen = 2, 2
	}
	for i := 0; i < width; i++ {
		r.drawDashedHLine(rect.Min.X, rect.Max.X, rect.Min.Y+i, c, dashLen, gapLen)
		r.drawDashedHLine(rect.Min.X, rect.Max.X, rect.Max.Y-1-i, c, dashLen, gapLen)
		r.drawDashedVLine(rect.Min.X+i, rect.Min.Y, rect.Max.Y, c, dashLen, gapLen)
		r.drawDashedVLine(rect.Max.X-1-i, rect.Min.Y, rect.Max.Y, c, dashLen, gapLen)
	}
}

func (r *renderer) drawDashedHLine(x1, x2, y int, c color.RGBA, dashLen, gapLen int) {
	period := dashLen + gapLen
	for x := x1; x < x2; x++ {
		if (x-x1)%period < dashLen {
			r.blendPixel(x, y, c)
		}
	}
}

func (r *renderer) drawDashedVLine(x, y1, y2 int, c color.RGBA, dashLen, gapLen int) {
	period := dashLen + gapLen
	for y := y1; y < y2; y++ {
		if (y-y1)%period < dashLen {
			r.blendPixel(x, y, c)
		}
	}
}

func (r *renderer) drawLineThick(x1, y1, x2, y2 int, c color.RGBA, width int) {
	if width <= 1 {
		r.drawLine(x1, y1, x2, y2, c)
		return
	}
	dx := float64(x2 - x1)
	dy := float64(y2 - y1)
	length := math.Sqrt(dx*dx + dy*dy)
	if length < 0.5 {
		r.blendPixel(x1, y1, c)
		return
	}
	nx := -dy / length
	ny := dx / length
	hw := float64(width) / 2.0
	for i := 0; i < width; i++ {
		offset := -hw + float64(i) + 0.5
		r.drawLine(x1+int(offset*nx), y1+int(offset*ny), x2+int(offset*nx), y2+int(offset*ny), c)
	}
}

func (r *renderer) drawLine(x1, y1, x2, y2 int, c color.RGBA) {
	dx := abs(x2 - x1)
	dy := abs(y2 - y1)
	sx, sy := 1, 1
	if x1 > x2 {
		sx = -1
	}
	if y1 > y2 {
		sy = -1
	}
	err := dx - dy
	for {
		r.blendPixel(x1, y1, c)
		if x1 == x2 && y1 == y2 {
			break
		}
		e2 := 2 * err
		if e2 > -dy {
			err -= dy
			x1 += sx
		}
		if e2 < dx {
			err += dx
			y1 += sy
		}
	}
}

func (r *renderer) drawLineAA(x1, y1, x2, y2 int, c color.RGBA, width int) {
	if width <= 1 {
		r.drawLineWu(float64(x1), float64(y1), float64(x2), float64(y2), c)
		return
	}
	dx := float64(x2 - x1)
	dy := float64(y2 - y1)
	length := math.Sqrt(dx*dx + dy*dy)
	if length < 0.5 {
		r.blendPixel(x1, y1, c)
		return
	}
	nx := -dy / length
	ny := dx / length
	hw := float64(width) / 2.0
	for i := 0; i < width; i++ {
		offset := -hw + float64(i) + 0.5
		ox := offset * nx
		oy := offset * ny
		r.drawLineWu(float64(x1)+ox, float64(y1)+oy, float64(x2)+ox, float64(y2)+oy, c)
	}
}

func (r *renderer) drawLineWu(x0, y0, x1, y1 float64, c color.RGBA) {
	steep := math.Abs(y1-y0) > math.Abs(x1-x0)
	if steep {
		x0, y0 = y0, x0
		x1, y1 = y1, x1
	}
	if x0 > x1 {
		x0, x1 = x1, x0
		y0, y1 = y1, y0
	}
	dx := x1 - x0
	dy := y1 - y0
	gradient := 0.0
	if dx != 0 {
		gradient = dy / dx
	}

	// First endpoint
	xend := math.Round(x0)
	yend := y0 + gradient*(xend-x0)
	xgap := 1.0 - fpart(x0+0.5)
	xpxl1 := int(xend)
	ypxl1 := int(math.Floor(yend))
	if steep {
		r.blendPixelF(ypxl1, xpxl1, c, (1-fpart(yend))*xgap)
		r.blendPixelF(ypxl1+1, xpxl1, c, fpart(yend)*xgap)
	} else {
		r.blendPixelF(xpxl1, ypxl1, c, (1-fpart(yend))*xgap)
		r.blendPixelF(xpxl1, ypxl1+1, c, fpart(yend)*xgap)
	}
	intery := yend + gradient

	// Second endpoint
	xend = math.Round(x1)
	yend = y1 + gradient*(xend-x1)
	xgap = fpart(x1 + 0.5)
	xpxl2 := int(xend)
	ypxl2 := int(math.Floor(yend))
	if steep {
		r.blendPixelF(ypxl2, xpxl2, c, (1-fpart(yend))*xgap)
		r.blendPixelF(ypxl2+1, xpxl2, c, fpart(yend)*xgap)
	} else {
		r.blendPixelF(xpxl2, ypxl2, c, (1-fpart(yend))*xgap)
		r.blendPixelF(xpxl2, ypxl2+1, c, fpart(yend)*xgap)
	}

	for x := xpxl1 + 1; x < xpxl2; x++ {
		iy := int(math.Floor(intery))
		f := fpart(intery)
		if steep {
			r.blendPixelF(iy, x, c, 1-f)
			r.blendPixelF(iy+1, x, c, f)
		} else {
			r.blendPixelF(x, iy, c, 1-f)
			r.blendPixelF(x, iy+1, c, f)
		}
		intery += gradient
	}
}

func fpart(x float64) float64 { return x - math.Floor(x) }

// --- Ellipse rendering (anti-aliased) ---

func (r *renderer) fillEllipseAA(cx, cy, w, h int, c color.RGBA) {
	if w <= 0 || h <= 0 {
		return
	}
	rx := float64(w) / 2
	ry := float64(h) / 2
	centerX := float64(cx) + rx
	centerY := float64(cy) + ry
	invRx2 := 1.0 / (rx * rx)
	invRy2 := 1.0 / (ry * ry)
	aaThreshold := 0.05

	bounds := r.img.Bounds()
	pix := r.img.Pix
	stride := r.img.Stride

	for py := cy; py < cy+h; py++ {
		if py < bounds.Min.Y || py >= bounds.Max.Y {
			continue
		}
		dyNorm := float64(py) + 0.5 - centerY
		dy2 := dyNorm * dyNorm * invRy2
		if dy2 > 1.0 {
			continue
		}
		hExtent := rx * math.Sqrt(1.0-dy2)
		minPx := maxInt(int(centerX-hExtent), cx)
		maxPx := minInt(int(centerX+hExtent+1), cx+w)
		minPx = maxInt(minPx, bounds.Min.X)
		maxPx = minInt(maxPx, bounds.Max.X)

		rowOff := (py-bounds.Min.Y)*stride + (minPx-bounds.Min.X)*4
		for px := minPx; px < maxPx; px++ {
			dxNorm := float64(px) + 0.5 - centerX
			d := dxNorm*dxNorm*invRx2 + dy2
			if d <= 1.0 {
				edge := 1.0 - d
				if edge < aaThreshold {
					r.blendPixelF(px, py, c, edge/aaThreshold)
				} else if c.A == 255 {
					pix[rowOff] = c.R
					pix[rowOff+1] = c.G
					pix[rowOff+2] = c.B
					pix[rowOff+3] = 255
				} else {
					a := uint32(c.A)
					ia := 255 - a
					pix[rowOff] = uint8((uint32(c.R)*a + uint32(pix[rowOff])*ia) / 255)
					pix[rowOff+1] = uint8((uint32(c.G)*a + uint32(pix[rowOff+1])*ia) / 255)
					pix[rowOff+2] = uint8((uint32(c.B)*a + uint32(pix[rowOff+2])*ia) / 255)
					pix[rowOff+3] = uint8(uint32(pix[rowOff+3]) + (255-uint32(pix[rowOff+3]))*a/255)
				}
			}
			rowOff += 4
		}
	}

}

func (r *renderer) drawEllipseAA(cx, cy, w, h int, c color.RGBA, lineWidth int) {
	if w <= 0 || h <= 0 {
		return
	}
	rx := float64(w) / 2
	ry := float64(h) / 2
	centerX := float64(cx) + rx
	centerY := float64(cy) + ry
	lw := float64(lineWidth)
	minR := math.Min(rx, ry)
	if minR < 1 {
		minR = 1
	}
	halfLW := lw / 2
	threshold := halfLW + 1

	for py := cy - lineWidth - 1; py < cy+h+lineWidth+1; py++ {
		dyNorm := (float64(py) + 0.5 - centerY) / ry
		dy2 := dyNorm * dyNorm
		if dy2 > 1.5 { // quick reject for rows far outside
			continue
		}
		for px := cx - lineWidth - 1; px < cx+w+lineWidth+1; px++ {
			dxNorm := (float64(px) + 0.5 - centerX) / rx
			d := math.Sqrt(dxNorm*dxNorm + dy2)
			distPx := math.Abs(d-1.0) * minR
			if distPx < threshold {
				coverage := 1.0
				if distPx > halfLW {
					coverage = 1.0 - (distPx - halfLW)
				}
				if coverage > 0 {
					r.blendPixelF(px, py, c, coverage)
				}
			}
		}
	}
}

// Legacy compatibility wrappers
func (r *renderer) fillEllipse(cx, cy, w, h int, c color.RGBA) { r.fillEllipseAA(cx, cy, w, h, c) }
func (r *renderer) drawEllipse(cx, cy, w, h int, c color.RGBA) { r.drawEllipseAA(cx, cy, w, h, c, 1) }

// --- Rounded rectangle ---

func (r *renderer) fillRoundedRect(x, y, w, h, radius int, c color.RGBA) {
	if radius <= 0 {
		r.fillRectBlend(image.Rect(x, y, x+w, y+h), c)
		return
	}
	radius = minInt(radius, minInt(w/2, h/2))
	r2 := float64(radius * radius)

	// Fill center rectangle (no corner checks needed)
	r.fillRectBlend(image.Rect(x+radius, y, x+w-radius, y+h), c)
	// Fill left/right strips (excluding corners)
	r.fillRectBlend(image.Rect(x, y+radius, x+radius, y+h-radius), c)
	r.fillRectBlend(image.Rect(x+w-radius, y+radius, x+w, y+h-radius), c)

	// Fill corners with circle test
	corners := [4][2]int{
		{x + radius, y + radius},         // top-left center
		{x + w - radius, y + radius},     // top-right center
		{x + radius, y + h - radius},     // bottom-left center
		{x + w - radius, y + h - radius}, // bottom-right center
	}
	cornerRects := [4]image.Rectangle{
		{Min: image.Pt(x, y), Max: image.Pt(x+radius, y+radius)},
		{Min: image.Pt(x+w-radius, y), Max: image.Pt(x+w, y+radius)},
		{Min: image.Pt(x, y+h-radius), Max: image.Pt(x+radius, y+h)},
		{Min: image.Pt(x+w-radius, y+h-radius), Max: image.Pt(x+w, y+h)},
	}
	for ci := 0; ci < 4; ci++ {
		ccx, ccy := corners[ci][0], corners[ci][1]
		cr := cornerRects[ci]
		for py := cr.Min.Y; py < cr.Max.Y; py++ {
			dy := float64(py - ccy)
			for px := cr.Min.X; px < cr.Max.X; px++ {
				dx := float64(px - ccx)
				if dx*dx+dy*dy <= r2 {
					r.blendPixel(px, py, c)
				}
			}
		}
	}
}

func (r *renderer) drawRoundedRect(x, y, w, h, radius int, c color.RGBA, lineWidth int) {
	r.drawLineThick(x+radius, y, x+w-radius, y, c, lineWidth)
	r.drawLineThick(x+radius, y+h-1, x+w-radius, y+h-1, c, lineWidth)
	r.drawLineThick(x, y+radius, x, y+h-radius, c, lineWidth)
	r.drawLineThick(x+w-1, y+radius, x+w-1, y+h-radius, c, lineWidth)
	r.drawArc(x, y, radius*2, radius*2, c, math.Pi, 1.5*math.Pi, lineWidth)
	r.drawArc(x+w-radius*2, y, radius*2, radius*2, c, 1.5*math.Pi, 2*math.Pi, lineWidth)
	r.drawArc(x, y+h-radius*2, radius*2, radius*2, c, 0.5*math.Pi, math.Pi, lineWidth)
	r.drawArc(x+w-radius*2, y+h-radius*2, radius*2, radius*2, c, 0, 0.5*math.Pi, lineWidth)
}

func (r *renderer) drawArc(cx, cy, w, h int, c color.RGBA, startAngle, endAngle float64, lineWidth int) {
	rx := float64(w) / 2
	ry := float64(h) / 2
	centerX := float64(cx) + rx
	centerY := float64(cy) + ry
	// Use enough steps for smooth arc
	circumference := math.Pi * (rx + ry) * (endAngle - startAngle) / (2 * math.Pi)
	steps := maxInt(int(circumference*2), 30)
	angleStep := (endAngle - startAngle) / float64(steps)

	var prevPx, prevPy int
	for i := 0; i <= steps; i++ {
		angle := startAngle + angleStep*float64(i)
		px := int(centerX + rx*math.Cos(angle))
		py := int(centerY + ry*math.Sin(angle))
		if i > 0 && (px != prevPx || py != prevPy) {
			r.drawLineThick(prevPx, prevPy, px, py, c, lineWidth)
		}
		prevPx, prevPy = px, py
	}
}

// --- Polygon shapes ---

type fpoint struct{ x, y float64 }

// fillPolygon fills a polygon using scanline algorithm with sort.Float64s.
func (r *renderer) fillPolygon(pts []fpoint, c color.RGBA) {
	if len(pts) < 3 {
		return
	}
	minY, maxY := pts[0].y, pts[0].y
	for _, p := range pts[1:] {
		if p.y < minY {
			minY = p.y
		}
		if p.y > maxY {
			maxY = p.y
		}
	}

	n := len(pts)
	// Pre-allocate intersection buffer
	intersections := make([]float64, 0, n)

	for y := int(minY); y <= int(maxY); y++ {
		fy := float64(y) + 0.5
		intersections = intersections[:0]
		for i := 0; i < n; i++ {
			j := (i + 1) % n
			py1, py2 := pts[i].y, pts[j].y
			if py1 > py2 {
				py1, py2 = py2, py1
			}
			if fy < py1 || fy >= py2 {
				continue
			}
			dy := pts[j].y - pts[i].y
			if dy == 0 {
				continue
			}
			t := (fy - pts[i].y) / dy
			intersections = append(intersections, pts[i].x+t*(pts[j].x-pts[i].x))
		}
		sort.Float64s(intersections)
		for i := 0; i+1 < len(intersections); i += 2 {
			x1 := int(math.Ceil(intersections[i]))
			x2 := int(math.Floor(intersections[i+1]))
			if x1 <= x2 {
				if c.A == 255 {
					r.fillRectFast(image.Rect(x1, y, x2+1, y+1), c)
				} else {
					r.fillRectBlend(image.Rect(x1, y, x2+1, y+1), c)
				}
			}
		}
	}
}

func (r *renderer) drawPolygon(pts []fpoint, c color.RGBA, width int) {
	n := len(pts)
	for i := 0; i < n; i++ {
		j := (i + 1) % n
		r.drawLineAA(int(pts[i].x), int(pts[i].y), int(pts[j].x), int(pts[j].y), c, width)
	}
}

func (r *renderer) fillTriangle(x, y, w, h int, c color.RGBA) {
	r.fillPolygon([]fpoint{
		{float64(x) + float64(w)/2, float64(y)},
		{float64(x + w), float64(y + h)},
		{float64(x), float64(y + h)},
	}, c)
}

func (r *renderer) drawTriangle(x, y, w, h int, c color.RGBA, width int) {
	r.drawPolygon([]fpoint{
		{float64(x) + float64(w)/2, float64(y)},
		{float64(x + w), float64(y + h)},
		{float64(x), float64(y + h)},
	}, c, width)
}

func (r *renderer) fillDiamond(x, y, w, h int, c color.RGBA) {
	cx, cy := float64(x)+float64(w)/2, float64(y)+float64(h)/2
	r.fillPolygon([]fpoint{{cx, float64(y)}, {float64(x + w), cy}, {cx, float64(y + h)}, {float64(x), cy}}, c)
}

func (r *renderer) drawDiamond(x, y, w, h int, c color.RGBA, width int) {
	cx, cy := float64(x)+float64(w)/2, float64(y)+float64(h)/2
	r.drawPolygon([]fpoint{{cx, float64(y)}, {float64(x + w), cy}, {cx, float64(y + h)}, {float64(x), cy}}, c, width)
}

func (r *renderer) fillRegularPolygon(x, y, w, h, sides int, startAngle float64, c color.RGBA) {
	cx := float64(x) + float64(w)/2
	cy := float64(y) + float64(h)/2
	rx := float64(w) / 2
	ry := float64(h) / 2
	pts := make([]fpoint, sides)
	for i := 0; i < sides; i++ {
		angle := startAngle + float64(i)*2*math.Pi/float64(sides)
		pts[i] = fpoint{cx + rx*math.Cos(angle), cy + ry*math.Sin(angle)}
	}
	r.fillPolygon(pts, c)
}

func (r *renderer) fillPentagon(x, y, w, h int, c color.RGBA) {
	r.fillRegularPolygon(x, y, w, h, 5, -math.Pi/2, c)
}

func (r *renderer) fillHexagon(x, y, w, h int, c color.RGBA) {
	r.fillRegularPolygon(x, y, w, h, 6, 0, c)
}

func (r *renderer) fillStar(x, y, w, h, points int, c color.RGBA) {
	cx := float64(x) + float64(w)/2
	cy := float64(y) + float64(h)/2
	outerRx, outerRy := float64(w)/2, float64(h)/2
	innerRx, innerRy := outerRx*0.4, outerRy*0.4
	n := points * 2
	pts := make([]fpoint, n)
	for i := 0; i < n; i++ {
		angle := -math.Pi/2 + float64(i)*2*math.Pi/float64(n)
		rx, ry := outerRx, outerRy
		if i%2 == 1 {
			rx, ry = innerRx, innerRy
		}
		pts[i] = fpoint{cx + rx*math.Cos(angle), cy + ry*math.Sin(angle)}
	}
	r.fillPolygon(pts, c)
}

func (r *renderer) fillArrowRight(x, y, w, h int, c color.RGBA) {
	shaftH := float64(h) * 0.4
	headW := float64(w) * 0.35
	shaftW := float64(w) - headW
	top := float64(y) + (float64(h)-shaftH)/2
	bot := top + shaftH
	r.fillPolygon([]fpoint{
		{float64(x), top}, {float64(x) + shaftW, top}, {float64(x) + shaftW, float64(y)},
		{float64(x + w), float64(y) + float64(h)/2},
		{float64(x) + shaftW, float64(y + h)}, {float64(x) + shaftW, bot}, {float64(x), bot},
	}, c)
}

func (r *renderer) fillArrowLeft(x, y, w, h int, c color.RGBA) {
	shaftH := float64(h) * 0.4
	headW := float64(w) * 0.35
	top := float64(y) + (float64(h)-shaftH)/2
	bot := top + shaftH
	r.fillPolygon([]fpoint{
		{float64(x + w), top}, {float64(x) + headW, top}, {float64(x) + headW, float64(y)},
		{float64(x), float64(y) + float64(h)/2},
		{float64(x) + headW, float64(y + h)}, {float64(x) + headW, bot}, {float64(x + w), bot},
	}, c)
}

func (r *renderer) fillArrowUp(x, y, w, h int, c color.RGBA) {
	shaftW := float64(w) * 0.4
	headH := float64(h) * 0.35
	left := float64(x) + (float64(w)-shaftW)/2
	right := left + shaftW
	r.fillPolygon([]fpoint{
		{float64(x) + float64(w)/2, float64(y)},
		{float64(x + w), float64(y) + headH}, {right, float64(y) + headH},
		{right, float64(y + h)}, {left, float64(y + h)},
		{left, float64(y) + headH}, {float64(x), float64(y) + headH},
	}, c)
}

func (r *renderer) fillArrowDown(x, y, w, h int, c color.RGBA) {
	shaftW := float64(w) * 0.4
	headH := float64(h) * 0.35
	shaftTop := float64(h) - headH
	left := float64(x) + (float64(w)-shaftW)/2
	right := left + shaftW
	r.fillPolygon([]fpoint{
		{left, float64(y)}, {right, float64(y)},
		{right, float64(y) + shaftTop}, {float64(x + w), float64(y) + shaftTop},
		{float64(x) + float64(w)/2, float64(y + h)},
		{float64(x), float64(y) + shaftTop}, {left, float64(y) + shaftTop},
	}, c)
}

func (r *renderer) fillHeart(x, y, w, h int, c color.RGBA) {
	cx := float64(x) + float64(w)/2
	topY := float64(y) + float64(h)*0.3
	halfW := float64(w) / 2
	hScale := float64(h) * 0.7

	for py := y; py < y+h; py++ {
		ny := 1 - (float64(py)-topY)/hScale
		ny2 := ny * ny
		ny3 := ny2 * ny
		for px := x; px < x+w; px++ {
			nx := (float64(px) - cx) / halfW
			nx2 := nx * nx
			val := (nx2 + ny2 - 1)
			val = val * val * val
			val -= nx2 * ny3
			if val <= 0 {
				r.blendPixel(px, py, c)
			}
		}
	}
}

func (r *renderer) fillPlus(x, y, w, h int, c color.RGBA) {
	armW := w / 3
	armH := h / 3
	r.fillRectBlend(image.Rect(x, y+armH, x+w, y+h-armH), c)
	r.fillRectBlend(image.Rect(x+armW, y, x+w-armW, y+h), c)
}

// --- Text rendering ---

// getFace returns a font.Face for the given Font, falling back to basicfont.Face7x13.
func (r *renderer) getFace(f *Font) font.Face {
	if r.fontCache == nil {
		return basicfont.Face7x13
	}
	sizePt := float64(f.Size)
	if sizePt <= 0 {
		sizePt = 10
	}
	// Scale font size by DPI ratio (fonts are defined at 72 DPI)
	sizePt = sizePt * r.dpi / 72.0

	face := r.fontCache.GetFace(f.Name, sizePt, f.Bold, f.Italic)
	if face != nil {
		return face
	}
	// CJK fallback names
	for _, fallback := range []string{
		"Microsoft YaHei", "SimSun", "SimHei", "NSimSun",
		"Yu Gothic", "Meiryo", "MS Gothic",
		"Malgun Gothic", "Gulim",
		"Noto Sans CJK SC", "Noto Sans SC", "WenQuanYi Micro Hei",
		"Arial", "Helvetica", "DejaVu Sans",
	} {
		face = r.fontCache.GetFace(fallback, sizePt, f.Bold, f.Italic)
		if face != nil {
			return face
		}
	}
	return basicfont.Face7x13
}

// textRun holds a measured run of text with its formatting.
type textRun struct {
	text  string
	font  *Font
	face  font.Face
	width int
}

// textLine holds a line of text runs with total metrics.
type textLine struct {
	runs       []textRun
	width      int
	ascent     int
	descent    int
	lineHeight int
}

// buildTextLine measures a slice of textRuns and returns a textLine.
func (r *renderer) buildTextLine(runs []textRun) textLine {
	var tl textLine
	tl.runs = runs
	for _, run := range runs {
		tl.width += run.width
		if run.face == nil {
			continue
		}
		metrics := run.face.Metrics()
		asc := metrics.Ascent.Ceil()
		desc := metrics.Descent.Ceil()
		if asc > tl.ascent {
			tl.ascent = asc
		}
		if desc > tl.descent {
			tl.descent = desc
		}
	}
	tl.lineHeight = tl.ascent + tl.descent
	if tl.lineHeight < 1 {
		tl.lineHeight = 14
	}
	return tl
}

// drawParagraphs renders paragraphs within the given bounding box.
func (r *renderer) drawParagraphs(paragraphs []*Paragraph, x, y, w, h int, anchor TextAnchorType) {
	if len(paragraphs) == 0 {
		return
	}

	// Build all lines from all paragraphs, tracking per-paragraph spacing
	type lineInfo struct {
		line        textLine
		spaceBefore int
		spaceAfter  int
		lineSpacing int // 0 means default (single)
		hAlign      HorizontalAlignment
		paraIdx     int  // index into paragraphs slice
		isFirst     bool // first line of paragraph
		isLast      bool // last line of paragraph
	}
	var allLines []lineInfo

	for pi, para := range paragraphs {
		align := HorizontalLeft
		marginLeft := 0
		marginRight := 0
		indent := 0
		if para.alignment != nil {
			align = para.alignment.Horizontal
			marginLeft = r.emuToPixelX(para.alignment.MarginLeft)
			marginRight = r.emuToPixelX(para.alignment.MarginRight)
			indent = r.emuToPixelX(para.alignment.Indent)
		}

		// Build runs for this paragraph
		var paraRuns []textRun

		// Bullet run
		if para.bullet != nil && para.bullet.Type != BulletTypeNone {
			bRun := r.buildBulletRun(para.bullet, para)
			if bRun.text != "" {
				paraRuns = append(paraRuns, bRun)
			}
		}

		for _, elem := range para.elements {
			switch e := elem.(type) {
			case *TextRun:
				if e.text == "" {
					continue
				}
				f := e.font
				if f == nil {
					f = NewFont()
				}
				face := r.getFace(f)
				paraRuns = append(paraRuns, textRun{
					text:  e.text,
					font:  f,
					face:  face,
					width: font.MeasureString(face, e.text).Ceil(),
				})
			case *BreakElement:
				// Force a new line
				paraRuns = append(paraRuns, textRun{text: "\n"})
			}
		}

		// Wrap runs into lines
		availW := w - marginLeft - marginRight - indent
		if availW < 10 {
			availW = w
		}
		lines := r.wrapRunLine(paraRuns, availW)
		if len(lines) == 0 {
			// Empty paragraph still takes space
			lines = []textLine{{lineHeight: 14}}
		}

		for i, line := range lines {
			li := lineInfo{
				line:        line,
				lineSpacing: para.lineSpacing,
				hAlign:      align,
				paraIdx:     pi,
				isFirst:     i == 0,
				isLast:      i == len(lines)-1,
			}
			if i == 0 {
				li.spaceBefore = r.emuToPixelY(int64(para.spaceBefore))
			}
			if i == len(lines)-1 {
				li.spaceAfter = r.emuToPixelY(int64(para.spaceAfter))
			}
			allLines = append(allLines, li)
		}
	}

	// Calculate total height
	totalH := 0
	for i, li := range allLines {
		if i > 0 {
			totalH += li.spaceBefore
		}
		lh := li.line.lineHeight
		if li.lineSpacing > 0 {
			lh = int(float64(lh) * float64(li.lineSpacing) / 10000.0)
		}
		totalH += lh
		totalH += li.spaceAfter
	}

	// Vertical anchor offset
	startY := y
	switch anchor {
	case TextAnchorMiddle:
		startY = y + (h-totalH)/2
	case TextAnchorBottom:
		startY = y + h - totalH
	}

	curY := startY
	for i, li := range allLines {
		if i > 0 {
			curY += li.spaceBefore
		}

		lh := li.line.lineHeight
		if li.lineSpacing > 0 {
			lh = int(float64(lh) * float64(li.lineSpacing) / 10000.0)
		}

		// Horizontal alignment
		lineX := x
		para := paragraphs[li.paraIdx]
		if para.alignment != nil {
			lineX += r.emuToPixelX(para.alignment.MarginLeft)
			if li.isFirst {
				lineX += r.emuToPixelX(para.alignment.Indent)
			}
		}

		switch li.hAlign {
		case HorizontalCenter:
			lineX = x + (w-li.line.width)/2
		case HorizontalRight:
			lineX = x + w - li.line.width
			if para.alignment != nil {
				lineX -= r.emuToPixelX(para.alignment.MarginRight)
			}
		}

		baseline := curY + li.line.ascent

		// Draw each run
		drawX := lineX
		for _, run := range li.line.runs {
			if run.text == "\n" || run.text == "" {
				continue
			}
			if run.face == nil {
				continue
			}
			fc := color.RGBA{A: 255}
			if run.font != nil {
				fc = argbToRGBA(run.font.Color)
			}

			runBaseline := baseline
			if run.font != nil {
				if run.font.Superscript {
					runBaseline -= li.line.ascent / 3
				} else if run.font.Subscript {
					runBaseline += li.line.descent / 2
				}
			}

			d := &font.Drawer{
				Dst:  r.img,
				Src:  image.NewUniform(fc),
				Face: run.face,
				Dot:  fixed.P(drawX, runBaseline),
			}
			d.DrawString(run.text)

			// Underline
			if run.font != nil && run.font.Underline != UnderlineNone {
				uy := runBaseline + 2
				r.drawUnderline(drawX, drawX+run.width, uy, fc, run.font.Underline)
			}

			// Strikethrough
			if run.font != nil && run.font.Strikethrough {
				sy := runBaseline - li.line.ascent/3
				r.drawLine(drawX, sy, drawX+run.width, sy, fc)
			}

			drawX += run.width
		}

		curY += lh
		curY += li.spaceAfter
	}
}

// drawUnderline draws an underline of the given style.
func (r *renderer) drawUnderline(x1, x2, y int, c color.RGBA, style UnderlineType) {
	switch style {
	case UnderlineSingle:
		r.drawLine(x1, y, x2, y, c)
	case UnderlineDouble:
		r.drawLine(x1, y-1, x2, y-1, c)
		r.drawLine(x1, y+1, x2, y+1, c)
	case UnderlineHeavy:
		r.drawLine(x1, y-1, x2, y-1, c)
		r.drawLine(x1, y, x2, y, c)
		r.drawLine(x1, y+1, x2, y+1, c)
	case UnderlineDash:
		r.drawDashedHLine(x1, x2, y, c, 6, 3)
	case UnderlineWavy:
		for px := x1; px < x2; px++ {
			wy := y + int(math.Sin(float64(px-x1)*0.5)*2)
			r.blendPixel(px, wy, c)
		}
	default:
		r.drawLine(x1, y, x2, y, c)
	}
}

// buildBulletRun creates a textRun for a bullet prefix.
func (r *renderer) buildBulletRun(b *Bullet, para *Paragraph) textRun {
	if b == nil || b.Type == BulletTypeNone {
		return textRun{}
	}

	// Determine bullet font
	bulletFont := NewFont()
	bulletFont.Size = 10
	// Try to get size from first text run
	for _, elem := range para.elements {
		if tr, ok := elem.(*TextRun); ok && tr.font != nil {
			bulletFont.Size = tr.font.Size
			break
		}
	}
	if b.Color != nil {
		bulletFont.Color = *b.Color
	}
	if b.Font != "" {
		bulletFont.Name = b.Font
	}

	var text string
	switch b.Type {
	case BulletTypeChar:
		text = b.Style + " "
	case BulletTypeNumeric, BulletTypeAutoNum:
		num := b.StartAt
		if num < 1 {
			num = 1
		}
		text = formatBulletNumber(num, b.NumFormat) + " "
	}

	face := r.getFace(bulletFont)
	return textRun{
		text:  text,
		font:  bulletFont,
		face:  face,
		width: font.MeasureString(face, text).Ceil(),
	}
}

// formatBulletNumber formats a number according to the bullet format.
func formatBulletNumber(num int, format string) string {
	switch format {
	case NumFormatRomanUcPeriod:
		return toRoman(num) + "."
	case NumFormatRomanLcPeriod:
		return strings.ToLower(toRoman(num)) + "."
	case NumFormatAlphaUcPeriod:
		if num >= 1 && num <= 26 {
			return string(rune('A'+num-1)) + "."
		}
		return fmt.Sprintf("%d.", num)
	case NumFormatAlphaLcPeriod:
		if num >= 1 && num <= 26 {
			return string(rune('a'+num-1)) + "."
		}
		return fmt.Sprintf("%d.", num)
	case NumFormatAlphaLcParen:
		if num >= 1 && num <= 26 {
			return string(rune('a'+num-1)) + ")"
		}
		return fmt.Sprintf("%d)", num)
	case NumFormatArabicParen:
		return fmt.Sprintf("%d)", num)
	default: // arabicPeriod
		return fmt.Sprintf("%d.", num)
	}
}

// toRoman converts an integer to a Roman numeral string.
func toRoman(num int) string {
	if num <= 0 || num > 3999 {
		return fmt.Sprintf("%d", num)
	}
	vals := []int{1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1}
	syms := []string{"M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"}
	var buf strings.Builder
	for i, v := range vals {
		for num >= v {
			buf.WriteString(syms[i])
			num -= v
		}
	}
	return buf.String()
}

// wrapRunLine wraps text runs into multiple lines that fit within maxWidth.
func (r *renderer) wrapRunLine(runs []textRun, maxWidth int) []textLine {
	if len(runs) == 0 {
		return nil
	}
	if maxWidth <= 0 {
		maxWidth = 1
	}

	var lines []textLine
	var currentRuns []textRun
	currentWidth := 0

	for _, run := range runs {
		if run.text == "\n" {
			lines = append(lines, r.buildTextLine(currentRuns))
			currentRuns = nil
			currentWidth = 0
			continue
		}
		if run.face == nil {
			continue
		}

		// If the run fits, add it
		if currentWidth+run.width <= maxWidth || len(currentRuns) == 0 {
			currentRuns = append(currentRuns, run)
			currentWidth += run.width
			continue
		}

		// Need to wrap: try word-level splitting
		words := strings.Fields(run.text)
		if len(words) <= 1 {
			// Single word doesn't fit, force it on current line or new line
			if len(currentRuns) > 0 {
				lines = append(lines, r.buildTextLine(currentRuns))
				currentRuns = nil
				currentWidth = 0
			}
			currentRuns = append(currentRuns, run)
			currentWidth = run.width
			continue
		}

		// Split by words
		var partial strings.Builder
		for i, word := range words {
			test := partial.String()
			if i > 0 {
				test += " "
			}
			test += word
			tw := font.MeasureString(run.face, test).Ceil()
			if currentWidth+tw > maxWidth && (len(currentRuns) > 0 || partial.Len() > 0) {
				if partial.Len() > 0 {
					pText := partial.String()
					currentRuns = append(currentRuns, textRun{
						text:  pText,
						font:  run.font,
						face:  run.face,
						width: font.MeasureString(run.face, pText).Ceil(),
					})
				}
				lines = append(lines, r.buildTextLine(currentRuns))
				currentRuns = nil
				currentWidth = 0
				partial.Reset()
				partial.WriteString(word)
			} else {
				if partial.Len() > 0 {
					partial.WriteString(" ")
				}
				partial.WriteString(word)
			}
		}
		if partial.Len() > 0 {
			pText := partial.String()
			wr := textRun{
				text:  pText,
				font:  run.font,
				face:  run.face,
				width: font.MeasureString(run.face, pText).Ceil(),
			}
			currentRuns = append(currentRuns, wr)
			currentWidth += wr.width
		}
	}

	if len(currentRuns) > 0 {
		lines = append(lines, r.buildTextLine(currentRuns))
	}

	return lines
}

// drawStringCentered draws a string centered in the given rectangle.
func (r *renderer) drawStringCentered(text string, face font.Face, c color.RGBA, rect image.Rectangle) {
	if text == "" || face == nil {
		return
	}
	tw := font.MeasureString(face, text).Ceil()
	metrics := face.Metrics()
	th := (metrics.Ascent + metrics.Descent).Ceil()
	cx := rect.Min.X + (rect.Dx()-tw)/2
	cy := rect.Min.Y + (rect.Dy()-th)/2 + metrics.Ascent.Ceil()
	d := &font.Drawer{
		Dst:  r.img,
		Src:  image.NewUniform(c),
		Face: face,
		Dot:  fixed.P(cx, cy),
	}
	d.DrawString(text)
}

// --- Chart rendering ---

// defaultChartPalette is the default color palette for chart series.
var defaultChartPalette = []color.RGBA{
	{R: 79, G: 129, B: 189, A: 255},
	{R: 192, G: 80, B: 77, A: 255},
	{R: 155, G: 187, B: 89, A: 255},
	{R: 128, G: 100, B: 162, A: 255},
	{R: 75, G: 172, B: 198, A: 255},
	{R: 247, G: 150, B: 70, A: 255},
	{R: 119, G: 44, B: 42, A: 255},
	{R: 77, G: 93, B: 58, A: 255},
}

// chartColors returns the default color palette for chart series.
func chartColors() []color.RGBA {
	return defaultChartPalette
}

// getSeriesColor returns the color for a series, using its FillColor if set, otherwise a palette color.
func getSeriesColor(s *ChartSeries, idx int, palette []color.RGBA) color.RGBA {
	if s.FillColor.ARGB != "" && s.FillColor.ARGB != "00000000" {
		return argbToRGBA(s.FillColor)
	}
	return palette[idx%len(palette)]
}

func (r *renderer) renderChart(s *ChartShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)

	// Background
	r.fillRectFast(image.Rect(x, y, x+w, y+h), color.RGBA{R: 255, G: 255, B: 255, A: 255})
	r.drawRect(image.Rect(x, y, x+w, y+h), color.RGBA{R: 200, G: 200, B: 200, A: 255}, 1)

	// Title
	titleH := 0
	if s.title != nil && s.title.Visible && s.title.Text != "" {
		face := r.getFace(s.title.Font)
		fc := argbToRGBA(s.title.Font.Color)
		titleH = face.Metrics().Height.Ceil() + 4
		r.drawStringCentered(s.title.Text, face, fc, image.Rect(x, y, x+w, y+titleH))
	}

	// Legend height
	legendH := 0
	if s.legend != nil && s.legend.Visible {
		legendH = 20
	}

	// Plot area
	plotX := x + 40
	plotY := y + titleH + 5
	plotW := w - 50
	plotH := h - titleH - legendH - 15
	if plotW < 10 {
		plotW = 10
	}
	if plotH < 10 {
		plotH = 10
	}

	ct := s.plotArea.GetType()
	if ct == nil {
		return
	}

	switch c := ct.(type) {
	case *BarChart:
		r.renderBarChart(c, plotX, plotY, plotW, plotH)
	case *Bar3DChart:
		r.renderBarChart(&c.BarChart, plotX, plotY, plotW, plotH)
	case *LineChart:
		r.renderLineChart(c, plotX, plotY, plotW, plotH)
	case *PieChart:
		r.renderPieChart(c.Series, plotX, plotY, plotW, plotH)
	case *Pie3DChart:
		r.renderPieChart(c.Series, plotX, plotY, plotW, plotH)
	case *DoughnutChart:
		r.renderDoughnutChart(c, plotX, plotY, plotW, plotH)
	case *AreaChart:
		r.renderAreaChart(c, plotX, plotY, plotW, plotH)
	case *ScatterChart:
		r.renderScatterChart(c, plotX, plotY, plotW, plotH)
	case *RadarChart:
		r.renderRadarChart(c, plotX, plotY, plotW, plotH)
	}

	// Legend
	if s.legend != nil && s.legend.Visible {
		r.renderChartLegend(s, x, y+h-legendH, w, legendH)
	}
}

func (r *renderer) renderBarChart(c *BarChart, px, py, pw, ph int) {
	if len(c.Series) == 0 {
		return
	}
	palette := chartColors()

	// Collect all categories and find value range
	cats := c.Series[0].Categories
	minVal := 0.0
	maxVal := 0.0
	first := true
	for _, s := range c.Series {
		for _, cat := range s.Categories {
			v := s.Values[cat]
			if first {
				minVal = v
				maxVal = v
				first = false
			} else {
				if v < minVal {
					minVal = v
				}
				if v > maxVal {
					maxVal = v
				}
			}
		}
	}
	if minVal > 0 {
		minVal = 0
	}
	if maxVal <= minVal {
		maxVal = minVal + 1
	}
	valRange := maxVal - minVal

	// Draw axes
	axisColor := color.RGBA{R: 128, G: 128, B: 128, A: 255}
	r.drawLine(px, py+ph, px+pw, py+ph, axisColor)
	r.drawLine(px, py, px, py+ph, axisColor)

	nCats := len(cats)
	nSeries := len(c.Series)
	if nCats == 0 {
		return
	}
	catW := pw / nCats
	barW := catW / (nSeries + 1)
	if barW < 1 {
		barW = 1
	}

	for ci, cat := range cats {
		for si, s := range c.Series {
			v := s.Values[cat]
			barH := int(float64(ph) * (v - minVal) / valRange)
			bx := px + ci*catW + (si+1)*barW - barW/2
			by := py + ph - barH
			sc := getSeriesColor(s, si, palette)
			r.fillRectBlend(image.Rect(bx, by, bx+barW-1, py+ph), sc)
		}
	}
}

func (r *renderer) renderLineChart(c *LineChart, px, py, pw, ph int) {
	if len(c.Series) == 0 {
		return
	}
	palette := chartColors()

	// Find value range
	minVal := math.MaxFloat64
	maxVal := -math.MaxFloat64
	for _, s := range c.Series {
		for _, v := range s.Values {
			if v < minVal {
				minVal = v
			}
			if v > maxVal {
				maxVal = v
			}
		}
	}
	if minVal > 0 {
		minVal = 0
	}
	if maxVal <= minVal {
		maxVal = minVal + 1
	}
	valRange := maxVal - minVal

	// Draw axes
	axisColor := color.RGBA{R: 128, G: 128, B: 128, A: 255}
	r.drawLine(px, py+ph, px+pw, py+ph, axisColor)
	r.drawLine(px, py, px, py+ph, axisColor)

	for si, s := range c.Series {
		sc := getSeriesColor(s, si, palette)
		cats := s.Categories
		nPts := len(cats)
		if nPts == 0 {
			continue
		}
		prevX, prevY := 0, 0
		for i, cat := range cats {
			v := s.Values[cat]
			ptX := px
			if nPts > 1 {
				ptX = px + i*pw/(nPts-1)
			}
			ptY := py + ph - int(float64(ph)*(v-minVal)/valRange)
			if i > 0 {
				r.drawLineAA(prevX, prevY, ptX, ptY, sc, 2)
			}
			// Draw marker
			r.fillEllipseAA(ptX-2, ptY-2, 5, 5, sc)
			prevX, prevY = ptX, ptY
		}
	}
}

func (r *renderer) renderPieChart(series []*ChartSeries, px, py, pw, ph int) {
	if len(series) == 0 || len(series[0].Categories) == 0 {
		return
	}
	palette := chartColors()
	s := series[0]

	// Sum values
	total := 0.0
	for _, cat := range s.Categories {
		v := s.Values[cat]
		if v > 0 {
			total += v
		}
	}
	if total == 0 {
		return
	}

	cx := px + pw/2
	cy := py + ph/2
	radius := minInt(pw, ph) / 2
	if radius < 5 {
		return
	}

	startAngle := -math.Pi / 2
	for i, cat := range s.Categories {
		v := s.Values[cat]
		if v <= 0 {
			continue
		}
		sweep := 2 * math.Pi * v / total
		endAngle := startAngle + sweep
		sc := palette[i%len(palette)]
		r.fillPieSlice(cx, cy, radius, startAngle, endAngle, sc)
		startAngle = endAngle
	}
}

// fillPieSlice fills a pie slice using scanline approach with row-level x-range.
func (r *renderer) fillPieSlice(cx, cy, radius int, startAngle, endAngle float64, c color.RGBA) {
	r2 := radius * radius
	for dy := -radius; dy <= radius; dy++ {
		dy2 := dy * dy
		if dy2 > r2 {
			continue
		}
		// Compute max dx for this row
		maxDx := int(math.Sqrt(float64(r2 - dy2)))
		for dx := -maxDx; dx <= maxDx; dx++ {
			angle := math.Atan2(float64(dy), float64(dx))
			if angleInSweep(angle, startAngle, endAngle) {
				r.blendPixel(cx+dx, cy+dy, c)
			}
		}
	}
}

// angleInSweep checks if angle is within the sweep from start to end (going clockwise).
func angleInSweep(angle, start, end float64) bool {
	// Normalize to [0, 2*pi)
	norm := func(a float64) float64 {
		for a < 0 {
			a += 2 * math.Pi
		}
		for a >= 2*math.Pi {
			a -= 2 * math.Pi
		}
		return a
	}
	a := norm(angle)
	s := norm(start)
	e := norm(end)
	if s <= e {
		return a >= s && a <= e
	}
	return a >= s || a <= e
}

func (r *renderer) renderDoughnutChart(c *DoughnutChart, px, py, pw, ph int) {
	if len(c.Series) == 0 || len(c.Series[0].Categories) == 0 {
		return
	}
	palette := chartColors()
	s := c.Series[0]

	total := 0.0
	for _, cat := range s.Categories {
		v := s.Values[cat]
		if v > 0 {
			total += v
		}
	}
	if total == 0 {
		return
	}

	cx := px + pw/2
	cy := py + ph/2
	outerR := minInt(pw, ph) / 2
	innerR := outerR * c.HoleSize / 100
	if outerR < 5 {
		return
	}

	startAngle := -math.Pi / 2
	for i, cat := range s.Categories {
		v := s.Values[cat]
		if v <= 0 {
			continue
		}
		sweep := 2 * math.Pi * v / total
		endAngle := startAngle + sweep
		sc := palette[i%len(palette)]
		r.fillDoughnutSlice(cx, cy, innerR, outerR, startAngle, endAngle, sc)
		startAngle = endAngle
	}
}

// fillDoughnutSlice fills a doughnut slice.
func (r *renderer) fillDoughnutSlice(cx, cy, innerR, outerR int, startAngle, endAngle float64, c color.RGBA) {
	or2 := outerR * outerR
	ir2 := innerR * innerR
	for dy := -outerR; dy <= outerR; dy++ {
		dy2 := dy * dy
		if dy2 > or2 {
			continue
		}
		maxDx := int(math.Sqrt(float64(or2 - dy2)))
		for dx := -maxDx; dx <= maxDx; dx++ {
			d2 := dx*dx + dy2
			if d2 < ir2 {
				continue
			}
			angle := math.Atan2(float64(dy), float64(dx))
			if angleInSweep(angle, startAngle, endAngle) {
				r.blendPixel(cx+dx, cy+dy, c)
			}
		}
	}
}

func (r *renderer) renderAreaChart(c *AreaChart, px, py, pw, ph int) {
	if len(c.Series) == 0 {
		return
	}
	palette := chartColors()

	minVal := math.MaxFloat64
	maxVal := -math.MaxFloat64
	for _, s := range c.Series {
		for _, v := range s.Values {
			if v < minVal {
				minVal = v
			}
			if v > maxVal {
				maxVal = v
			}
		}
	}
	if minVal > 0 {
		minVal = 0
	}
	if maxVal <= minVal {
		maxVal = minVal + 1
	}
	valRange := maxVal - minVal

	// Axes
	axisColor := color.RGBA{R: 128, G: 128, B: 128, A: 255}
	r.drawLine(px, py+ph, px+pw, py+ph, axisColor)
	r.drawLine(px, py, px, py+ph, axisColor)

	for si, s := range c.Series {
		sc := getSeriesColor(s, si, palette)
		// Semi-transparent fill
		fillC := color.RGBA{R: sc.R, G: sc.G, B: sc.B, A: 128}
		cats := s.Categories
		nPts := len(cats)
		if nPts == 0 {
			continue
		}

		pts := make([]fpoint, 0, nPts+2)
		for i, cat := range cats {
			v := s.Values[cat]
			ptX := float64(px)
			if nPts > 1 {
				ptX = float64(px) + float64(i)*float64(pw)/float64(nPts-1)
			}
			ptY := float64(py+ph) - float64(ph)*(v-minVal)/valRange
			pts = append(pts, fpoint{ptX, ptY})
		}
		// Close polygon along baseline
		pts = append(pts, fpoint{pts[len(pts)-1].x, float64(py + ph)})
		pts = append(pts, fpoint{pts[0].x, float64(py + ph)})
		r.fillPolygon(pts, fillC)

		// Draw line on top
		for i := 0; i < nPts-1; i++ {
			r.drawLineAA(int(pts[i].x), int(pts[i].y), int(pts[i+1].x), int(pts[i+1].y), sc, 2)
		}
	}
}

func (r *renderer) renderScatterChart(c *ScatterChart, px, py, pw, ph int) {
	if len(c.Series) == 0 {
		return
	}
	palette := chartColors()

	// For scatter, categories are X values (parsed as indices), values are Y
	minVal := math.MaxFloat64
	maxVal := -math.MaxFloat64
	for _, s := range c.Series {
		for _, v := range s.Values {
			if v < minVal {
				minVal = v
			}
			if v > maxVal {
				maxVal = v
			}
		}
	}
	if minVal > 0 {
		minVal = 0
	}
	if maxVal <= minVal {
		maxVal = minVal + 1
	}
	valRange := maxVal - minVal

	axisColor := color.RGBA{R: 128, G: 128, B: 128, A: 255}
	r.drawLine(px, py+ph, px+pw, py+ph, axisColor)
	r.drawLine(px, py, px, py+ph, axisColor)

	for si, s := range c.Series {
		sc := getSeriesColor(s, si, palette)
		cats := s.Categories
		nPts := len(cats)
		if nPts == 0 {
			continue
		}
		for i, cat := range cats {
			v := s.Values[cat]
			ptX := px + (i * pw / maxInt(nPts-1, 1))
			ptY := py + ph - int(float64(ph)*(v-minVal)/valRange)
			r.fillEllipseAA(ptX-3, ptY-3, 7, 7, sc)
		}
	}
}

func (r *renderer) renderRadarChart(c *RadarChart, px, py, pw, ph int) {
	if len(c.Series) == 0 {
		return
	}
	palette := chartColors()

	// Find max value
	maxVal := 0.0
	for _, s := range c.Series {
		for _, v := range s.Values {
			if v > maxVal {
				maxVal = v
			}
		}
	}
	if maxVal == 0 {
		maxVal = 1
	}

	cx := px + pw/2
	cy := py + ph/2
	radius := minInt(pw, ph) / 2

	// Draw radar grid
	gridColor := color.RGBA{R: 200, G: 200, B: 200, A: 255}
	nCats := len(c.Series[0].Categories)
	if nCats == 0 {
		return
	}
	for i := 0; i < nCats; i++ {
		angle := 2*math.Pi*float64(i)/float64(nCats) - math.Pi/2
		ex := cx + int(float64(radius)*math.Cos(angle))
		ey := cy + int(float64(radius)*math.Sin(angle))
		r.drawLine(cx, cy, ex, ey, gridColor)
	}

	// Draw series
	for si, s := range c.Series {
		sc := getSeriesColor(s, si, palette)
		cats := s.Categories
		nPts := len(cats)
		if nPts == 0 {
			continue
		}
		pts := make([]fpoint, nPts)
		for i, cat := range cats {
			v := s.Values[cat]
			angle := 2*math.Pi*float64(i)/float64(nPts) - math.Pi/2
			dist := float64(radius) * v / maxVal
			pts[i] = fpoint{
				x: float64(cx) + dist*math.Cos(angle),
				y: float64(cy) + dist*math.Sin(angle),
			}
		}
		// Draw polygon
		for i := 0; i < nPts; i++ {
			j := (i + 1) % nPts
			r.drawLineAA(int(pts[i].x), int(pts[i].y), int(pts[j].x), int(pts[j].y), sc, 2)
		}
		// Fill with semi-transparent
		fillC := color.RGBA{R: sc.R, G: sc.G, B: sc.B, A: 64}
		r.fillPolygon(pts, fillC)
	}
}

func (r *renderer) renderChartLegend(s *ChartShape, lx, ly, lw, lh int) {
	ct := s.plotArea.GetType()
	if ct == nil {
		return
	}
	palette := chartColors()
	face := r.getFace(s.legend.Font)

	var names []string
	var colors []color.RGBA

	switch c := ct.(type) {
	case *BarChart:
		for i, ser := range c.Series {
			names = append(names, ser.Title)
			colors = append(colors, getSeriesColor(ser, i, palette))
		}
	case *Bar3DChart:
		for i, ser := range c.Series {
			names = append(names, ser.Title)
			colors = append(colors, getSeriesColor(ser, i, palette))
		}
	case *LineChart:
		for i, ser := range c.Series {
			names = append(names, ser.Title)
			colors = append(colors, getSeriesColor(ser, i, palette))
		}
	case *PieChart:
		if len(c.Series) > 0 {
			for i, cat := range c.Series[0].Categories {
				names = append(names, cat)
				colors = append(colors, palette[i%len(palette)])
			}
		}
	case *Pie3DChart:
		if len(c.Series) > 0 {
			for i, cat := range c.Series[0].Categories {
				names = append(names, cat)
				colors = append(colors, palette[i%len(palette)])
			}
		}
	case *DoughnutChart:
		if len(c.Series) > 0 {
			for i, cat := range c.Series[0].Categories {
				names = append(names, cat)
				colors = append(colors, palette[i%len(palette)])
			}
		}
	case *AreaChart:
		for i, ser := range c.Series {
			names = append(names, ser.Title)
			colors = append(colors, getSeriesColor(ser, i, palette))
		}
	case *ScatterChart:
		for i, ser := range c.Series {
			names = append(names, ser.Title)
			colors = append(colors, getSeriesColor(ser, i, palette))
		}
	case *RadarChart:
		for i, ser := range c.Series {
			names = append(names, ser.Title)
			colors = append(colors, getSeriesColor(ser, i, palette))
		}
	}

	if len(names) == 0 {
		return
	}

	// Draw legend entries horizontally centered
	entryW := lw / len(names)
	for i, name := range names {
		ex := lx + i*entryW
		// Color box
		boxSize := 10
		bx := ex + 4
		by := ly + (lh-boxSize)/2
		r.fillRectFast(image.Rect(bx, by, bx+boxSize, by+boxSize), colors[i])
		// Text
		d := &font.Drawer{
			Dst:  r.img,
			Src:  image.NewUniform(color.RGBA{A: 255}),
			Face: face,
			Dot:  fixed.P(bx+boxSize+4, ly+lh/2+4),
		}
		d.DrawString(name)
	}
}

// --- Image scaling ---

// scaleImageBilinear scales an image to the target width and height using bilinear interpolation.
func scaleImageBilinear(src image.Image, dstW, dstH int) *image.RGBA {
	if dstW <= 0 || dstH <= 0 {
		return image.NewRGBA(image.Rect(0, 0, 1, 1))
	}
	bounds := src.Bounds()
	srcW := bounds.Dx()
	srcH := bounds.Dy()
	if srcW <= 0 || srcH <= 0 {
		return image.NewRGBA(image.Rect(0, 0, dstW, dstH))
	}

	dst := image.NewRGBA(image.Rect(0, 0, dstW, dstH))

	xRatio := float64(srcW) / float64(dstW)
	yRatio := float64(srcH) / float64(dstH)

	// Fast path for *image.RGBA source
	if srcRGBA, ok := src.(*image.RGBA); ok {
		for dy := 0; dy < dstH; dy++ {
			sy := float64(dy) * yRatio
			sy0 := int(sy)
			sy1 := sy0 + 1
			if sy1 >= srcH {
				sy1 = srcH - 1
			}
			fy := sy - float64(sy0)
			ify := 1 - fy
			srcOff0 := (sy0+bounds.Min.Y-srcRGBA.Rect.Min.Y)*srcRGBA.Stride + (bounds.Min.X-srcRGBA.Rect.Min.X)*4
			srcOff1 := (sy1+bounds.Min.Y-srcRGBA.Rect.Min.Y)*srcRGBA.Stride + (bounds.Min.X-srcRGBA.Rect.Min.X)*4
			dstOff := dy * dst.Stride

			for dx := 0; dx < dstW; dx++ {
				sx := float64(dx) * xRatio
				sx0 := int(sx)
				sx1 := sx0 + 1
				if sx1 >= srcW {
					sx1 = srcW - 1
				}
				fx := sx - float64(sx0)
				ifx := 1 - fx

				o00 := srcOff0 + sx0*4
				o10 := srcOff0 + sx1*4
				o01 := srcOff1 + sx0*4
				o11 := srcOff1 + sx1*4
				sp := srcRGBA.Pix

				for ch := 0; ch < 4; ch++ {
					top := float64(sp[o00+ch])*ifx + float64(sp[o10+ch])*fx
					bot := float64(sp[o01+ch])*ifx + float64(sp[o11+ch])*fx
					dst.Pix[dstOff+ch] = uint8(top*ify + bot*fy)
				}
				dstOff += 4
			}
		}
		return dst
	}

	// Generic path for other image types
	for dy := 0; dy < dstH; dy++ {
		sy := float64(dy) * yRatio
		sy0 := int(sy)
		sy1 := sy0 + 1
		if sy1 >= srcH {
			sy1 = srcH - 1
		}
		fy := sy - float64(sy0)

		for dx := 0; dx < dstW; dx++ {
			sx := float64(dx) * xRatio
			sx0 := int(sx)
			sx1 := sx0 + 1
			if sx1 >= srcW {
				sx1 = srcW - 1
			}
			fx := sx - float64(sx0)

			r00, g00, b00, a00 := src.At(bounds.Min.X+sx0, bounds.Min.Y+sy0).RGBA()
			r10, g10, b10, a10 := src.At(bounds.Min.X+sx1, bounds.Min.Y+sy0).RGBA()
			r01, g01, b01, a01 := src.At(bounds.Min.X+sx0, bounds.Min.Y+sy1).RGBA()
			r11, g11, b11, a11 := src.At(bounds.Min.X+sx1, bounds.Min.Y+sy1).RGBA()

			lerp := func(v00, v10, v01, v11 uint32) uint8 {
				top := float64(v00)*(1-fx) + float64(v10)*fx
				bot := float64(v01)*(1-fx) + float64(v11)*fx
				return uint8((top*(1-fy) + bot*fy) / 256)
			}

			off := dy*dst.Stride + dx*4
			dst.Pix[off+0] = lerp(r00, r10, r01, r11)
			dst.Pix[off+1] = lerp(g00, g10, g01, g11)
			dst.Pix[off+2] = lerp(b00, b10, b01, b11)
			dst.Pix[off+3] = lerp(a00, a10, a01, a11)
		}
	}
	return dst
}

// scaleImage scales an image using nearest-neighbor (fast fallback).
func scaleImage(src image.Image, dstW, dstH int) *image.RGBA {
	return scaleImageBilinear(src, dstW, dstH)
}

// --- Utility functions ---

func abs(x int) int {
	if x < 0 {
		return -x
	}
	return x
}

func minInt(a, b int) int {
	if a < b {
		return a
	}
	return b
}

func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}
