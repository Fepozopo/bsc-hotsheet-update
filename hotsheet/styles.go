package hotsheet

import "github.com/xuri/excelize/v2"

const (
	// currencyFormat is the shared Excel number format used for dollar-value columns.
	currencyFormat = "$#,##0.00;[Red]($#,##0.00)"
	// Shared fill colors keep the workbook styling consistent across sheets.
	dataInsightsSectionFill = "#D9EAF7"
	standardHeaderFill      = "#E6E6FA"
	// dataInsightsTotalFill stays darker than the section/header fills so total rows remain the
	// strongest visual endpoint in each table.
	dataInsightsTotalFill = "#E2E2E2"
)

// currencyNumFmt returns a pointer suitable for excelize.Style.CustomNumFmt.
func currencyNumFmt() *string {
	format := currencyFormat
	return &format
}

// centeredAlignment returns the standard centered alignment fragment.
func centeredAlignment() *excelize.Alignment {
	return &excelize.Alignment{Horizontal: "center", Vertical: "center"}
}

// thinBlackBorder returns the standard four-sided black border fragment.
func thinBlackBorder() []excelize.Border {
	return []excelize.Border{
		{Type: "left", Color: "000000", Style: 1},
		{Type: "right", Color: "000000", Style: 1},
		{Type: "top", Color: "000000", Style: 1},
		{Type: "bottom", Color: "000000", Style: 1},
	}
}

// boldFont returns a bold font fragment, optionally with a size.
func boldFont(size ...float64) *excelize.Font {
	font := &excelize.Font{Bold: true}
	if len(size) > 0 {
		font.Size = size[0]
	}
	return font
}

// patternFill returns a patterned fill fragment for a solid background color.
func patternFill(fillColor string) excelize.Fill {
	return excelize.Fill{Type: "pattern", Color: []string{fillColor}, Pattern: 1}
}
