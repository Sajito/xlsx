package xlsx

import "encoding/xml"

type RelationshipType string

func (r RelationshipType) MarshalXMLAttr(name xml.Name) (xml.Attr, error) {
	if r == "" {
		return xml.Attr{}, nil
	}

	return xml.Attr{Name: name, Value: string(r)}, nil
}

const (
	RelationshipTypeHyperlink RelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)

type RelationshipTargetMode string

func (r RelationshipTargetMode) MarshalXMLAttr(name xml.Name) (xml.Attr, error) {
	if r == "" {
		return xml.Attr{}, nil
	}

	return xml.Attr{Name: name, Value: string(r)}, nil
}

const (
	RelationshipTargetModeExternal RelationshipTargetMode = "External"
)

type xlsxRels struct {
	XMLName       xml.Name       `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []xlsxRelation `xml:"Relationship"`
}

type xlsxRelation struct {
	Id         string                 `xml:"Id,attr"`
	Type       RelationshipType       `xml:"Type,attr"`
	Target     string                 `xml:"Target,attr"`
	TargetMode RelationshipTargetMode `xml:"TargetMode,attr,omitempty"`
}
