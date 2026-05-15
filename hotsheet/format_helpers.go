package hotsheet

import "github.com/xuri/excelize/v2"

const (
	// currencyFormat is the shared Excel number format used for dollar-value columns.
	currencyFormat = "$#,##0.00;[Red]($#,##0.00)"
	// Shared fill colors keep the workbook styling consistent across sheets.
	dataInsightsSectionFill = "#D9EAF7"
	standardHeaderFill      = "#E6E6FA"
	standardTotalFill       = "#F2F2F2"
)

// currencyNumFmt returns a pointer suitable for excelize.Style.CustomNumFmt.
func currencyNumFmt() *string {
	format := currencyFormat
	return &format
}

// centeredAlignmentStyle returns the standard centered alignment with no border.
func centeredAlignmentStyle() *excelize.Style {
	return &excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	}
}

// centeredBorderStyle returns the standard centered style with thin black borders.
func centeredBorderStyle() *excelize.Style {
	style := centeredAlignmentStyle()
	style.Border = []excelize.Border{
		{Type: "left", Color: "000000", Style: 1},
		{Type: "right", Color: "000000", Style: 1},
		{Type: "top", Color: "000000", Style: 1},
		{Type: "bottom", Color: "000000", Style: 1},
	}
	return style
}

// centeredFontStyle returns the standard centered style with the provided font.
func centeredFontStyle(font *excelize.Font) *excelize.Style {
	style := centeredAlignmentStyle()
	style.Font = font
	return style
}

// centeredFillStyle returns the standard centered style with a patterned fill.
func centeredFillStyle(fillColor string) *excelize.Style {
	style := centeredBorderStyle()
	style.Fill = excelize.Fill{Type: "pattern", Color: []string{fillColor}, Pattern: 1}
	return style
}

// centeredFillFontStyle returns the standard centered filled style with the provided font.
func centeredFillFontStyle(fillColor string, font *excelize.Font) *excelize.Style {
	style := centeredFillStyle(fillColor)
	style.Font = font
	return style
}

// centeredNumFmtStyle returns the standard centered style with the provided number format.
func centeredNumFmtStyle(numFmt *string) *excelize.Style {
	style := centeredBorderStyle()
	style.CustomNumFmt = numFmt
	return style
}

// centeredFillNumFmtStyle returns the standard centered filled style with the provided number format.
func centeredFillNumFmtStyle(fillColor string, numFmt *string) *excelize.Style {
	style := centeredFillStyle(fillColor)
	style.CustomNumFmt = numFmt
	return style
}

// centeredFillFontNumFmtStyle returns the standard centered filled style with the provided font and number format.
func centeredFillFontNumFmtStyle(fillColor string, font *excelize.Font, numFmt *string) *excelize.Style {
	style := centeredFillFontStyle(fillColor, font)
	style.CustomNumFmt = numFmt
	return style
}
