package gopresentation

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
	"time"
)

// --- Core Properties ---

func (r *PPTXReader) readCoreProperties(zr *zip.Reader, pres *Presentation) error {
	data, err := readFileFromZip(zr, "docProps/core.xml")
	if err != nil {
		return err
	}

	decoder := xml.NewDecoder(strings.NewReader(string(data)))
	props := pres.properties

	var currentElement string
	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			currentElement = t.Name.Local
		case xml.CharData:
			text := strings.TrimSpace(string(t))
			if text == "" {
				continue
			}
			switch currentElement {
			case "creator":
				props.Creator = text
			case "lastModifiedBy":
				props.LastModifiedBy = text
			case "title":
				props.Title = text
			case "description":
				props.Description = text
			case "subject":
				props.Subject = text
			case "keywords":
				props.Keywords = text
			case "category":
				props.Category = text
			case "revision":
				props.Revision = text
			case "created":
				if t, err := time.Parse("2006-01-02T15:04:05Z", text); err == nil {
					props.Created = t
				}
			case "modified":
				if t, err := time.Parse("2006-01-02T15:04:05Z", text); err == nil {
					props.Modified = t
				}
			}
		}
	}
	return nil
}

// --- Presentation ---

type xmlPresentation struct {
	XMLName        xml.Name           `xml:"presentation"`
	SldMasterIdLst xmlSldMasterIdLst  `xml:"sldMasterIdLst"`
	SldIdLst       xmlSldIdLst        `xml:"sldIdLst"`
	SldSz          xmlSldSz           `xml:"sldSz"`
	NotesSz        xmlNotesSz         `xml:"notesSz"`
}

type xmlSldMasterIdLst struct {
	SldMasterIds []xmlSldMasterId `xml:"sldMasterId"`
}

type xmlSldMasterId struct {
	ID  string `xml:"id,attr"`
}

type xmlSldIdLst struct {
	SldIds []xmlSldId `xml:"sldId"`
}

type xmlSldId struct {
	ID  string `xml:"id,attr"`
}

type xmlSldSz struct {
	CX   string `xml:"cx,attr"`
	CY   string `xml:"cy,attr"`
	Type string `xml:"type,attr"`
}

type xmlNotesSz struct {
	CX string `xml:"cx,attr"`
	CY string `xml:"cy,attr"`
}

func (r *PPTXReader) readPresentation(zr *zip.Reader, pres *Presentation) ([]string, error) {
	data, err := readFileFromZip(zr, "ppt/presentation.xml")
	if err != nil {
		return nil, fmt.Errorf("failed to read presentation.xml: %w", err)
	}

	// Parse using streaming to handle namespaces properly
	decoder := xml.NewDecoder(strings.NewReader(string(data)))
	var slideRelIDs []string

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "sldSz":
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "cx":
						if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							pres.layout.CX = v
						}
					case "cy":
						if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							pres.layout.CY = v
						}
					case "type":
						pres.layout.Name = attr.Value
					}
				}
			case "sldId":
				for _, attr := range t.Attr {
					if attr.Name.Local == "id" && attr.Name.Space != "" {
						slideRelIDs = append(slideRelIDs, attr.Value)
					} else if attr.Name.Local == "id" && attr.Name.Space == "" {
						// This is the numeric ID, not the relationship ID
					}
				}
			}
		}
	}

	// If we didn't find relationship IDs via namespace, try reading rels directly
	if len(slideRelIDs) == 0 {
		rels, err := r.readRelationships(zr, "ppt/_rels/presentation.xml.rels")
		if err == nil {
			for _, rel := range rels {
				if rel.Type == relTypeSlide {
					slideRelIDs = append(slideRelIDs, rel.ID)
				}
			}
		}
	}

	return slideRelIDs, nil
}
