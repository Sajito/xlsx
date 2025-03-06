package xlsx

import "encoding/xml"

type xlsxDrawing struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing wsDr"`

	TwoCellAnchors []xlsxTwoCellAnchor `xml:"twoCellAnchor"`
}

type xlsxTwoCellAnchor struct {
	EditAs string `xml:"editAs,attr"`

	From       xlsxMarker     `xml:"from"`
	To         xlsxMarker     `xml:"to"`
	Pic        *xlsxPicture   `xml:"pic"`
	ClientData xlsxClientData `xml:"clientData"`
}

type xlsxMarker struct {
	Col    int `xml:"col"`
	ColOff int `xml:"colOff"`
	Row    int `xml:"row"`
	RowOff int `xml:"rowOff"`
}

type xlsxPicture struct {
	Macro     string `xml:"macro,attr,omitempty"`
	Published bool   `xml:"published,attr,omitempty"`

	NvPicPr  xlsxPictureNonVisual `xml:"nvPicPr"`
	BlipFill xlsxBlipFill         `xml:"blipFill"`
	SpPr     *xlsxShapeProps      `xml:"spPr,omitempty"`
	// Style    *xlsxShapeStyle      `xml:"style,omitempty"` // - Not implemented
}

type xlsxPictureNonVisual struct {
	NvPr    xlsxNonVisualDrawingProps `xml:"cNvPr"`
	NvPicPr xlsxNonVisualPictureProps `xml:"cNvPicPr"`
}

type xlsxNonVisualPictureProps struct {
	PreferRelativeResize *bool `xml:"preferRelativeResize,attr,omitempty"` // default true

	PicLocks *xlsxPictureLocking
}

type xlsxPictureLocking struct {
	NoGrp              bool `xml:"noGrp,attr,omitempty"`
	NoSelect           bool `xml:"noSelect,attr,omitempty"`
	NoRot              bool `xml:"noRot,attr,omitempty"`
	NoChangeAspect     bool `xml:"noChangeAspect,attr,omitempty"`
	NoMove             bool `xml:"noMove,attr,omitempty"`
	NoResize           bool `xml:"noResize,attr,omitempty"`
	NoEditPoints       bool `xml:"noEditPoints,attr,omitempty"`
	NoAdjustHandles    bool `xml:"noAdjustHandles,attr,omitempty"`
	NoChangeArrowheads bool `xml:"noChangeArrowheads,attr,omitempty"`
	NoChangeShapeType  bool `xml:"noChangeShapeType,attr,omitempty"`

	NoCrop bool `xml:"noCrop,attr,omitempty"`
}

type xlsxNonVisualDrawingProps struct {
	Id     int    `xml:"id,attr"`
	Name   string `xml:"name,attr"`
	Desc   string `xml:"desc,attr,omitempty"`
	Hidden bool   `xml:"hidden,attr,omitempty"`
	Title  string `xml:"title,attr,omitempty"`

	HlinkClick *xlsxDrawingHyperlink `xml:"hlinkClick,omitempty"`
	HlinkHover *xlsxDrawingHyperlink `xml:"hlinkHover,omitempty"`
}

type xlsxDrawingHyperlink struct {
	Id             string `xml:"id,attr,omitempty"`
	InvalidUrl     string `xml:"invalidUrl,attr,omitempty"`
	Action         string `xml:"action,attr,omitempty"`
	TgtFarme       string `xml:"tgtFarme,attr,omitempty"`
	Tooltip        string `xml:"tooltip,attr,omitempty"`
	History        *bool  `xml:"history,attr,omitempty"` // default is true
	HighlightClick bool   `xml:"highlightClick,attr,omitempty"`
	EndSnd         bool   `xml:"endSnd,attr,omitempty"`
}

type xlsxBlipFill struct {
	Dpi          uint64 `xml:"dpi,attr,omitempty"`
	RotWithShape bool   `xml:"rotWithShape,attr,omitempty"`

	Blip *xlsxBlip `xml:"http://schemas.openxmlformats.org/drawingml/2006/main blip,omitempty"`
	// SrcRect - missing
	// Tile - missing
	Stretch xlsxStretch `xml:"http://schemas.openxmlformats.org/drawingml/2006/main stretch,omitempty"`
}

type xlsxBlip struct {
	Embed  string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships embed,attr"`
	Link   string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships link,attr,omitempty"`
	Cstate string `xml:"cstate,attr,omitempty"`
}

type xlsxStretch struct{}

type xlsxShapeProps struct {
	BwMode string `xml:"bwMode,attr,omitempty"`

	Xfrm     *xlsxTransform      `xml:"http://schemas.openxmlformats.org/drawingml/2006/main xfrm,omitempty"`
	PrstGeom *xlsxPresetGeometry `xml:"http://schemas.openxmlformats.org/drawingml/2006/main prstGeom,omitempty"`
	NoFill   *xlsxNoFill         `xml:"http://schemas.openxmlformats.org/drawingml/2006/main noFill,omitempty"`
	Ln       *xlsxLineProperties `xml:"http://schemas.openxmlformats.org/drawingml/2006/main ln,omitempty"`
}

type xlsxTransform struct {
	Rot   string `xml:"rot,attr,omitempty"`
	FlipH bool   `xml:"flipH,attr,omitempty"`
	FlipV bool   `xml:"flipV,attr,omitempty"`

	Off *xlsxPoint              `xml:"http://schemas.openxmlformats.org/drawingml/2006/main off,omitempty"`
	Ext *xlsxPositiveCoordinate `xml:"http://schemas.openxmlformats.org/drawingml/2006/main ext,omitempty"`
}

type xlsxPoint struct {
	X int `xml:"x,attr"`
	Y int `xml:"y,attr"`
}

type xlsxPositiveCoordinate struct {
	Cx uint `xml:"cx,attr"`
	Cy uint `xml:"cy,attr"`
}

type xlsxPresetGeometry struct {
	Prst string `xml:"prst,attr"`
}

type xlsxNoFill struct{}

type xlsxLineProperties struct {
	W    uint32 `xml:"w,attr"`
	Cap  string `xml:"cap,attr,omitempty"`
	Cmpd string `xml:"cmpd,attr,omitempty"`
	Algn string `xml:"algn,attr,omitempty"`
}

type xlsxClientData struct {
	LocksWithSheet  *bool `xml:"fLocksWithSheet,attr"`  // pointer, as default is true, therefore "false" must not be omitted
	PrintsWithSheet *bool `xml:"fPrintsWithSheet,attr"` // pointer, as default is true, therefore "false" must not be omitted
}
