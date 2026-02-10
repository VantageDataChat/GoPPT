package gopresentation

// EMU (English Metric Units) conversion helpers.
// 1 inch = 914400 EMU, 1 point = 12700 EMU, 1 cm = 360000 EMU.

// Inch converts inches to EMU.
func Inch(n float64) int64 {
	return int64(n * 914400)
}

// Point converts points to EMU.
func Point(n float64) int64 {
	return int64(n * 12700)
}

// Centimeter converts centimeters to EMU.
func Centimeter(n float64) int64 {
	return int64(n * 360000)
}

// Millimeter converts millimeters to EMU.
func Millimeter(n float64) int64 {
	return int64(n * 36000)
}

// EMUToInch converts EMU to inches.
func EMUToInch(emu int64) float64 {
	return float64(emu) / 914400
}

// EMUToPoint converts EMU to points.
func EMUToPoint(emu int64) float64 {
	return float64(emu) / 12700
}

// EMUToCentimeter converts EMU to centimeters.
func EMUToCentimeter(emu int64) float64 {
	return float64(emu) / 360000
}
