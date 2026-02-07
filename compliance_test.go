package gopresentation

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"strings"
	"testing"
)

// TestOOXMLCompliance validates that generated PPTX files conform to
// the Office Open XML (ECMA-376 / ISO 29500) standard structure.

func TestOOXMLContentTypesCompliance(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	slide.CreateRichTextShape().SetHeight(100).SetWidth(200)

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, err := readFileFromZip(zr, "[Content_Types].xml")
	if err != nil {
		t.Fatalf("missing [Content_Types].xml: %v", err)
	}

	// Must be valid XML
	var ct xmlContentTypes
	if err := xml.Unmarshal(data, &ct); err != nil {
		t.Fatalf("[Content_Types].xml is not valid XML: %v", err)
	}

	// Must have correct namespace
	if ct.Xmlns != nsContentTypes {
		t.Errorf("wrong namespace: %s", ct.Xmlns)
	}

	// Must have rels and xml defaults
	hasRels := false
	hasXML := false
	for _, d := range ct.Defaults {
		if d.Extension == "rels" {
			hasRels = true
		}
		if d.Extension == "xml" {
			hasXML = true
		}
	}
	if !hasRels {
		t.Error("missing rels default content type")
	}
	if !hasXML {
		t.Error("missing xml default content type")
	}

	// Must have required overrides
	requiredParts := map[string]bool{
		"/ppt/presentation.xml":                  false,
		"/ppt/presProps.xml":                     false,
		"/ppt/viewProps.xml":                     false,
		"/ppt/tableStyles.xml":                   false,
		"/ppt/slideMasters/slideMaster1.xml":     false,
		"/ppt/slideLayouts/slideLayout1.xml":      false,
		"/ppt/theme/theme1.xml":                  false,
		"/docProps/core.xml":                     false,
		"/docProps/app.xml":                      false,
		"/ppt/slides/slide1.xml":                 false,
	}

	for _, o := range ct.Overrides {
		if _, ok := requiredParts[o.PartName]; ok {
			requiredParts[o.PartName] = true
		}
	}

	for part, found := range requiredParts {
		if !found {
			t.Errorf("missing required content type override for: %s", part)
		}
	}
}

func TestOOXMLRootRelationshipsCompliance(t *testing.T) {
	p := New()

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, err := readFileFromZip(zr, "_rels/.rels")
	if err != nil {
		t.Fatalf("missing _rels/.rels: %v", err)
	}

	var rels xmlRelationships
	if err := xml.Unmarshal(data, &rels); err != nil {
		t.Fatalf("_rels/.rels is not valid XML: %v", err)
	}

	// Must reference presentation, core props, and extended props
	hasPresentation := false
	hasCoreProps := false
	hasExtProps := false

	for _, rel := range rels.Relationships {
		switch rel.Type {
		case relTypeOfficeDoc:
			hasPresentation = true
			if rel.Target != "ppt/presentation.xml" {
				t.Errorf("office document target should be ppt/presentation.xml, got %s", rel.Target)
			}
		case relTypeCoreProps:
			hasCoreProps = true
		case relTypeExtProps:
			hasExtProps = true
		}
	}

	if !hasPresentation {
		t.Error("root rels missing office document relationship")
	}
	if !hasCoreProps {
		t.Error("root rels missing core properties relationship")
	}
	if !hasExtProps {
		t.Error("root rels missing extended properties relationship")
	}
}

func TestOOXMLPresentationRelationshipsCompliance(t *testing.T) {
	p := New()
	p.CreateSlide() // 2 slides

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, err := readFileFromZip(zr, "ppt/_rels/presentation.xml.rels")
	if err != nil {
		t.Fatalf("missing presentation rels: %v", err)
	}

	var rels xmlRelationships
	if err := xml.Unmarshal(data, &rels); err != nil {
		t.Fatalf("presentation rels is not valid XML: %v", err)
	}

	hasSlideMaster := false
	slideCount := 0
	hasPresProps := false
	hasViewProps := false
	hasTableStyles := false
	hasTheme := false

	for _, rel := range rels.Relationships {
		switch rel.Type {
		case relTypeSlideMaster:
			hasSlideMaster = true
		case relTypeSlide:
			slideCount++
		case relTypePresProps:
			hasPresProps = true
		case relTypeViewProps:
			hasViewProps = true
		case relTypeTableStyles:
			hasTableStyles = true
		case relTypeTheme:
			hasTheme = true
		}
	}

	if !hasSlideMaster {
		t.Error("missing slide master relationship")
	}
	if slideCount != 2 {
		t.Errorf("expected 2 slide relationships, got %d", slideCount)
	}
	if !hasPresProps {
		t.Error("missing presentation properties relationship")
	}
	if !hasViewProps {
		t.Error("missing view properties relationship")
	}
	if !hasTableStyles {
		t.Error("missing table styles relationship")
	}
	if !hasTheme {
		t.Error("missing theme relationship")
	}
}

func TestOOXMLSlideStructureCompliance(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	rt := slide.CreateRichTextShape()
	rt.SetHeight(300).SetWidth(600).SetOffsetX(100).SetOffsetY(200)
	rt.CreateTextRun("Compliance test")

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/slide1.xml")
	content := string(data)

	// Slide must have p:sld root element
	if !strings.Contains(content, "<p:sld") {
		t.Error("slide must have p:sld root element")
	}

	// Must have correct namespaces
	if !strings.Contains(content, nsPresentationML) {
		t.Error("slide must reference PresentationML namespace")
	}
	if !strings.Contains(content, nsDrawingML) {
		t.Error("slide must reference DrawingML namespace")
	}
	if !strings.Contains(content, nsOfficeDocRels) {
		t.Error("slide must reference relationships namespace")
	}

	// Must have cSld element
	if !strings.Contains(content, "p:cSld") {
		t.Error("slide must have p:cSld element")
	}

	// Must have spTree (shape tree)
	if !strings.Contains(content, "p:spTree") {
		t.Error("slide must have p:spTree element")
	}

	// Must have nvGrpSpPr
	if !strings.Contains(content, "p:nvGrpSpPr") {
		t.Error("slide must have p:nvGrpSpPr element")
	}

	// Must have grpSpPr
	if !strings.Contains(content, "p:grpSpPr") {
		t.Error("slide must have p:grpSpPr element")
	}

	// Must have clrMapOvr
	if !strings.Contains(content, "p:clrMapOvr") {
		t.Error("slide must have p:clrMapOvr element")
	}
}

func TestOOXMLSlideMasterCompliance(t *testing.T) {
	p := New()

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slideMasters/slideMaster1.xml")
	content := string(data)

	// Must have sldMaster root
	if !strings.Contains(content, "p:sldMaster") {
		t.Error("slide master must have p:sldMaster root")
	}

	// Must have clrMap
	if !strings.Contains(content, "p:clrMap") {
		t.Error("slide master must have p:clrMap")
	}

	// Must have sldLayoutIdLst
	if !strings.Contains(content, "p:sldLayoutIdLst") {
		t.Error("slide master must have p:sldLayoutIdLst")
	}

	// Must reference at least one slide layout
	if !strings.Contains(content, "p:sldLayoutId") {
		t.Error("slide master must reference at least one slide layout")
	}

	// Slide master rels must reference slide layout and theme
	relsData, _ := readFileFromZip(zr, "ppt/slideMasters/_rels/slideMaster1.xml.rels")
	relsContent := string(relsData)

	if !strings.Contains(relsContent, relTypeSlideLayout) {
		t.Error("slide master rels must reference slide layout")
	}
	if !strings.Contains(relsContent, relTypeTheme) {
		t.Error("slide master rels must reference theme")
	}
}

func TestOOXMLSlideLayoutCompliance(t *testing.T) {
	p := New()

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slideLayouts/slideLayout1.xml")
	content := string(data)

	// Must have sldLayout root
	if !strings.Contains(content, "p:sldLayout") {
		t.Error("slide layout must have p:sldLayout root")
	}

	// Must have cSld
	if !strings.Contains(content, "p:cSld") {
		t.Error("slide layout must have p:cSld")
	}

	// Must have clrMapOvr
	if !strings.Contains(content, "p:clrMapOvr") {
		t.Error("slide layout must have p:clrMapOvr")
	}

	// Slide layout rels must reference slide master
	relsData, _ := readFileFromZip(zr, "ppt/slideLayouts/_rels/slideLayout1.xml.rels")
	if !strings.Contains(string(relsData), relTypeSlideMaster) {
		t.Error("slide layout rels must reference slide master")
	}
}

func TestOOXMLThemeCompliance(t *testing.T) {
	p := New()

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/theme/theme1.xml")
	content := string(data)

	// Must have theme root
	if !strings.Contains(content, "a:theme") {
		t.Error("theme must have a:theme root")
	}

	// Must have themeElements
	if !strings.Contains(content, "a:themeElements") {
		t.Error("theme must have a:themeElements")
	}

	// Must have color scheme with all required colors
	requiredColors := []string{
		"a:dk1", "a:lt1", "a:dk2", "a:lt2",
		"a:accent1", "a:accent2", "a:accent3", "a:accent4", "a:accent5", "a:accent6",
		"a:hlink", "a:folHlink",
	}
	for _, color := range requiredColors {
		if !strings.Contains(content, color) {
			t.Errorf("theme must have color: %s", color)
		}
	}

	// Must have font scheme
	if !strings.Contains(content, "a:fontScheme") {
		t.Error("theme must have a:fontScheme")
	}
	if !strings.Contains(content, "a:majorFont") {
		t.Error("theme must have a:majorFont")
	}
	if !strings.Contains(content, "a:minorFont") {
		t.Error("theme must have a:minorFont")
	}

	// Must have format scheme
	if !strings.Contains(content, "a:fmtScheme") {
		t.Error("theme must have a:fmtScheme")
	}
}

func TestOOXMLPresentationPartCompliance(t *testing.T) {
	p := New()
	p.CreateSlide()

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/presentation.xml")
	content := string(data)

	// Must have presentation root
	if !strings.Contains(content, "p:presentation") {
		t.Error("must have p:presentation root")
	}

	// Must have sldMasterIdLst
	if !strings.Contains(content, "p:sldMasterIdLst") {
		t.Error("must have p:sldMasterIdLst")
	}

	// Must have sldIdLst
	if !strings.Contains(content, "p:sldIdLst") {
		t.Error("must have p:sldIdLst")
	}

	// Must have sldSz
	if !strings.Contains(content, "p:sldSz") {
		t.Error("must have p:sldSz")
	}

	// Must have notesSz
	if !strings.Contains(content, "p:notesSz") {
		t.Error("must have p:notesSz")
	}

	// Slide IDs must start at 256 (per OOXML spec)
	if !strings.Contains(content, `id="256"`) {
		t.Error("first slide ID should be 256")
	}
	if !strings.Contains(content, `id="257"`) {
		t.Error("second slide ID should be 257")
	}

	// Slide master ID must be >= 2147483648
	if !strings.Contains(content, `id="2147483648"`) {
		t.Error("slide master ID should be >= 2147483648")
	}
}

func TestOOXMLCorePropertiesCompliance(t *testing.T) {
	p := New()
	p.GetDocumentProperties().Creator = "Test"

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "docProps/core.xml")
	content := string(data)

	// Must have correct namespaces
	if !strings.Contains(content, nsCoreProperties) {
		t.Error("core properties must have correct namespace")
	}
	if !strings.Contains(content, nsDC) {
		t.Error("core properties must have dc namespace")
	}
	if !strings.Contains(content, nsDCTerms) {
		t.Error("core properties must have dcterms namespace")
	}

	// Must have created and modified dates in W3CDTF format
	if !strings.Contains(content, `xsi:type="dcterms:W3CDTF"`) {
		t.Error("dates must use W3CDTF format")
	}
}

func TestOOXMLSlideRelationshipsCompliance(t *testing.T) {
	p := New()

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	data, _ := readFileFromZip(zr, "ppt/slides/_rels/slide1.xml.rels")
	content := string(data)

	// Each slide must reference a slide layout
	if !strings.Contains(content, relTypeSlideLayout) {
		t.Error("slide rels must reference slide layout")
	}
}

func TestOOXMLAllXMLWellFormed(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// Add all shape types
	rt := slide.CreateRichTextShape()
	rt.SetHeight(100).SetWidth(200).SetOffsetX(10).SetOffsetY(10)
	tr := rt.CreateTextRun("Text with special chars: <>&\"'")
	tr.GetFont().SetBold(true).SetItalic(true).SetSize(20)

	as := slide.CreateAutoShape()
	as.SetAutoShapeType(AutoShapeEllipse)
	as.BaseShape.SetOffsetX(300).SetOffsetY(300).SetWidth(200).SetHeight(200)
	as.GetFill().SetSolid(ColorRed)
	as.SetText("Shape text")

	line := slide.CreateLineShape()
	line.BaseShape.SetOffsetX(0).SetOffsetY(500).SetWidth(5000000).SetHeight(0)

	table := slide.CreateTableShape(2, 2)
	table.SetWidth(4000000).SetHeight(1000000)
	table.GetCell(0, 0).SetText("A")
	table.GetCell(0, 1).SetText("B")
	table.GetCell(1, 0).SetText("C")
	table.GetCell(1, 1).SetText("D")

	img := slide.CreateDrawingShape()
	img.SetImageData(createMinimalPNG(), "image/png")
	img.SetHeight(50).SetWidth(50)

	p.CreateSlide()

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	err := w.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo error: %v", err)
	}

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	xmlFileCount := 0
	for _, f := range zr.File {
		if strings.HasSuffix(f.Name, ".xml") || strings.HasSuffix(f.Name, ".rels") {
			rc, err := f.Open()
			if err != nil {
				t.Errorf("failed to open %s: %v", f.Name, err)
				continue
			}

			decoder := xml.NewDecoder(rc)
			for {
				_, err := decoder.Token()
				if err != nil {
					if err.Error() == "EOF" {
						break
					}
					t.Errorf("malformed XML in %s: %v", f.Name, err)
					break
				}
			}
			rc.Close()
			xmlFileCount++
		}
	}

	if xmlFileCount < 10 {
		t.Errorf("expected at least 10 XML files, got %d", xmlFileCount)
	}
	t.Logf("validated %d XML files for well-formedness", xmlFileCount)
}

func TestOOXMLPackageStructure(t *testing.T) {
	p := New()

	var buf bytes.Buffer
	w, _ := NewWriter(p, WriterPowerPoint2007)
	w.WriteTo(&buf)

	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))

	// Verify the package structure follows OOXML conventions
	requiredDirs := map[string]bool{
		"_rels/":                          false,
		"docProps/":                       false,
		"ppt/":                            false,
		"ppt/_rels/":                      false,
		"ppt/slides/":                     false,
		"ppt/slides/_rels/":               false,
		"ppt/slideMasters/":               false,
		"ppt/slideMasters/_rels/":         false,
		"ppt/slideLayouts/":               false,
		"ppt/slideLayouts/_rels/":         false,
		"ppt/theme/":                      false,
	}

	for _, f := range zr.File {
		for dir := range requiredDirs {
			if strings.HasPrefix(f.Name, dir) {
				requiredDirs[dir] = true
			}
		}
	}

	for dir, found := range requiredDirs {
		if !found {
			t.Errorf("missing required directory structure: %s", dir)
		}
	}
}
