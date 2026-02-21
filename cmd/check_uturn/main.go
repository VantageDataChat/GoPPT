package main

import (
	"fmt"
	"image"
	"image/png"
	"os"
	"strings"
)

func main() {
	f, _ := os.Open("cmd/render_slide7/slide7.png")
	defer f.Close()
	img, _ := png.Decode(f)
	bounds := img.Bounds()
	_ = bounds

	minX, maxX, minY, maxY := 9999, 0, 9999, 0
	count := 0
	for y := 500; y < 950; y++ {
		for x := 1400; x < 1780; x++ {
			cr, cg, cb, ca := img.At(x, y).RGBA()
			cr, cg, cb, ca = cr>>8, cg>>8, cb>>8, ca>>8
			if cr > 150 && cg < 30 && cb < 30 && ca > 200 {
				count++
				if x < minX { minX = x }
				if x > maxX { maxX = x }
				if y < minY { minY = y }
				if y > maxY { maxY = y }
			}
		}
	}
	fmt.Printf("Red pixels: %d\n", count)
	fmt.Printf("BBox: x=[%d,%d] y=[%d,%d] size=%dx%d\n", minX, maxX, minY, maxY, maxX-minX+1, maxY-minY+1)

	// ASCII art - wider for better visibility
	cols := 80
	rows := 35
	w := maxX - minX + 1
	h := maxY - minY + 1
	fmt.Printf("\nASCII (%dx%d):\n", w, h)
	for row := 0; row < rows; row++ {
		cy := minY + row*h/rows
		line := strings.Builder{}
		for col := 0; col < cols; col++ {
			cx := minX + col*w/cols
			redCount := 0
			total := 0
			for dy := 0; dy < h/rows && dy < 3; dy++ {
				for dx := 0; dx < w/cols && dx < 3; dx++ {
					cr, cg, cb, ca := img.At(cx+dx, cy+dy).RGBA()
					cr, cg, cb, ca = cr>>8, cg>>8, cb>>8, ca>>8
					total++
					if cr > 150 && cg < 30 && cb < 30 && ca > 200 {
						redCount++
					}
				}
			}
			if redCount > total/2 {
				line.WriteByte('#')
			} else if redCount > 0 {
				line.WriteByte('.')
			} else {
				line.WriteByte(' ')
			}
		}
		fmt.Printf("  |%s|\n", line.String())
	}

	// Column density - count red pixels per column
	fmt.Println("\nColumn density (red pixels per x-column):")
	for cx := minX; cx <= maxX; cx += 10 {
		cnt := 0
		for cy := minY; cy <= maxY; cy++ {
			cr, cg, cb, ca := img.At(cx, cy).RGBA()
			cr, cg, cb, ca = cr>>8, cg>>8, cb>>8, ca>>8
			if cr > 150 && cg < 30 && cb < 30 && ca > 200 {
				cnt++
			}
		}
		fmt.Printf("  x=%d: %d px\n", cx, cnt)
	}
	_ = image.Black
}
