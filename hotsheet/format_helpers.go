package hotsheet

import "github.com/xuri/excelize/v2"

// currencyFormat is the shared Excel number format used for dollar-value columns.
const currencyFormat = "$#,##0.00;[Red]($#,##0.00)"

// currencyNumFmt returns a pointer suitable for excelize.Style.CustomNumFmt.
func currencyNumFmt() *string {
	format := currencyFormat
	return &format
}

// centeredBorderStyle returns the standard centered style with thin black borders.
func centeredBorderStyle() *excelize.Style {
	return &excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	}
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
