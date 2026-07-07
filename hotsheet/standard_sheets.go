package hotsheet

import (
	"fmt"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// standardSheetNames preserves the original workbook tab order used by the hotsheet.
var standardSheetNames = []string{"Everyday", "Winter", "Spring"}

// writeStandardSheets writes the Everyday, Winter, and Spring tabs, their headers, their rows,
// and the shared widths and filters used by the standard hotsheet layout.
func writeStandardSheets(f *excelize.File, entries []*inventoryEntry, hasPO bool) error {
	headers, mtoYtdIdx, mtoPyIdx := buildStandardSheetHeaders(hasPO)

	for _, sheetName := range standardSheetNames {
		if err := writeStandardSheetHeaders(f, sheetName, headers, hasPO); err != nil {
			return err
		}
	}

	monthsThrough := currentMonthsThrough(time.Now())
	for _, sheetName := range standardSheetNames {
		if err := writeStandardSheetRows(f, sheetName, entries, hasPO, monthsThrough, mtoYtdIdx, mtoPyIdx); err != nil {
			return err
		}
	}

	if err := applyStandardSheetWidths(f, headers); err != nil {
		return err
	}
	if err := applyStandardSheetFilters(f, headers); err != nil {
		return err
	}

	return nil
}

// buildStandardSheetHeaders returns the header row used by the three standard report sheets and
// the indexes of the MTO columns used for conditional formatting.
func buildStandardSheetHeaders(hasPO bool) ([]string, int, int) {
	headers := []string{"Item Code", "QTY on Hand"}
	if hasPO {
		headers = append(headers,
			"PO Num 1",
			"QTY on PO 1",
			"PO Num 2",
			"QTY on PO 2",
		)
	}
	headers = append(headers,
		"Total QTY on PO",
		"QTY on SO+BO",
		"QTY Available",
		"MTO YTD",
		"MTO PY",
		"QTY Sold+Issued YTD",
		"QTY Sold+Issued PY",
		"Class",
		"Status",
		"Occasion",
		"Description",
		"UPC",
		"Foil",
		"Royalty Code",
		"Dollar Sold YTD",
		"Dollar Sold PY",
	)

	mtoYtdIdx, mtoPyIdx := -1, -1
	for i, h := range headers {
		switch h {
		case "MTO YTD":
			mtoYtdIdx = i
		case "MTO PY":
			mtoPyIdx = i
		}
	}
	return headers, mtoYtdIdx, mtoPyIdx
}

// writeStandardSheetHeaders writes the standard header row, applies the existing header style,
// and keeps the explanatory MTO comments attached to the corresponding columns.
func writeStandardSheetHeaders(f *excelize.File, sheetName string, headers []string, hasPO bool) error {
	_ = hasPO // The header layout already captures whether PO columns should be present.

	headerStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(),
		Border:    thinBlackBorder(),
		Fill:      patternFill(standardHeaderFill),
		Font:      boldFont(),
	})
	if err != nil {
		return fmt.Errorf("failed to create standard header style: %w", err)
	}

	for c, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(c+1, 1)
		if err := f.SetCellValue(sheetName, cell, h); err != nil {
			return fmt.Errorf("failed to set header cell %s on %s: %w", cell, sheetName, err)
		}
		if err := f.SetCellStyle(sheetName, cell, cell, headerStyle); err != nil {
			return fmt.Errorf("failed to style header cell %s on %s: %w", cell, sheetName, err)
		}

		// Keep the original worksheet guidance available directly in the header row.
		if h == "MTO YTD" {
			cmt := excelize.Comment{
				Cell:   cell,
				Author: "Shane DuPrey",
				Text:   "MTO YTD = QTY Available / ((QTY Sold+Issued YTD + QTY on SO+BO) / (monthsThrough + 1)). monthsThrough is the number of months completed in the current year (fractional). This shows months till out using year-to-date sales pace including current sales orders/backorders.",
				Height: 190,
				Width:  200,
			}
			_ = f.AddComment(sheetName, cmt)
		}
		if h == "MTO PY" {
			cmt := excelize.Comment{
				Cell:   cell,
				Author: "Shane DuPrey",
				Text:   "MTO PY = QTY Available / ((QTY Sold+Issued PY) / (salesSeason + 1)). salesSeason used: Winter=6.5, Spring=5, Everyday=12. This shows months till out using prior-year sales scaled to the season length.",
				Height: 180,
				Width:  180,
			}
			_ = f.AddComment(sheetName, cmt)
		}
	}

	return nil
}

// writeStandardSheetRows writes the report rows for one standard worksheet, preserving the
// current derived values, class-prefix behavior, and conditional coloring rules.
func writeStandardSheetRows(f *excelize.File, sheetName string, entries []*inventoryEntry, hasPO bool, monthsThrough float64, mtoYtdIdx, mtoPyIdx int) error {
	rowIdx := 2
	for _, e := range entries {
		sh := mapOccasion(e.Occasion)
		if sh != sheetName {
			continue
		}

		// Determine the sales-season window used for MTO PY calculations.
		// Winter and Spring use their shorter merchandising seasons, while Everyday uses
		// the full year so the historical sales pace stays consistent with the workbook notes.
		salesSeason := 12.0
		switch sh {
		case "Winter":
			salesSeason = 6.5
		case "Spring":
			salesSeason = 5.0
		default:
			salesSeason = 12.0
		}

		// Calculate the derived values used by the standard report layout.
		onSOBO := e.OnSO + e.OnBO
		totalInventory := e.OnHand + e.OnPO
		totalAvail := totalInventory - onSOBO

		totalSoldYTD := e.YTDSold + max(e.YTDIssued, 0)
		totalSoldPY := e.SoldPY + max(e.IssuedPY, 0)
		soldPerMonthYTD := (float64(totalSoldYTD) + float64(onSOBO)) / monthsThrough
		soldPerMonthPY := float64(totalSoldPY) / salesSeason

		mtoYTD := float64(totalAvail) / (soldPerMonthYTD + 1)
		mtoPY := float64(totalAvail) / (soldPerMonthPY + 1)

		classDesc := applyStandardDisplayClassPrefix(e)

		vals := []interface{}{
			e.SKU,
			e.OnHand,
		}
		if hasPO {
			vals = append(vals, e.PONum1, e.OnPO1, e.PONum2, e.OnPO2)
		}
		vals = append(vals,
			e.OnPO,
			onSOBO,
			totalAvail,
			mtoYTD,
			mtoPY,
			totalSoldYTD,
			totalSoldPY,
			classDesc,
			e.Status,
			e.Occasion,
			e.Description,
			e.UPC,
			e.Foil,
			e.RoyaltyCode,
			e.DollarSoldYTD,
			e.DollarSoldPY,
		)

		dollarYTDCol := len(vals) - 2
		dollarPYCol := len(vals) - 1

		for c, v := range vals {
			cell, _ := excelize.CoordinatesToCellName(c+1, rowIdx)
			if err := f.SetCellValue(sheetName, cell, v); err != nil {
				return fmt.Errorf("failed to write %s cell %s: %w", sheetName, cell, err)
			}

			fillColor := standardSheetCellFillColor(e.Status, c, mtoYtdIdx, mtoPyIdx, mtoYTD, mtoPY, v)
			styleDef := &excelize.Style{
				Alignment: centeredAlignment(),
				Border:    thinBlackBorder(),
				Fill:      patternFill(fillColor),
			}
			if c == dollarYTDCol || c == dollarPYCol {
				styleDef.CustomNumFmt = currencyNumFmt()
			}
			style, err := f.NewStyle(styleDef)
			if err != nil {
				return fmt.Errorf("failed to create %s cell style for %s: %w", sheetName, cell, err)
			}
			if err := f.SetCellStyle(sheetName, cell, cell, style); err != nil {
				return fmt.Errorf("failed to style %s cell %s: %w", sheetName, cell, err)
			}
		}

		rowIdx++
	}

	return nil
}

// applyStandardDisplayClassPrefix applies the current display-time class prefix rules while keeping the
// original inventory class available through RawClassDesc for downstream reuse.
func applyStandardDisplayClassPrefix(e *inventoryEntry) string {
	classDesc := strings.TrimSpace(e.ClassDesc)
	skuUpper := strings.ToUpper(strings.TrimSpace(e.SKU))
	prefix := ""

	// Check longer or more specific suffixes first so the display prefix is deterministic.
	switch {
	case strings.HasSuffix(skuUpper, "-LLB") || strings.HasSuffix(skuUpper, "LLB"):
		prefix = "LLB - "
	case strings.HasSuffix(skuUpper, "-TB") || strings.HasSuffix(skuUpper, "TB") || strings.HasPrefix(skuUpper, "TB"):
		prefix = "TB - "
	case strings.HasSuffix(skuUpper, "-WM") || strings.HasSuffix(skuUpper, "WM"):
		prefix = "WM - "
	case strings.HasSuffix(skuUpper, "-AN") || strings.HasSuffix(skuUpper, "AN"):
		prefix = "AN - "
	case strings.HasSuffix(skuUpper, "-BN") || strings.HasSuffix(skuUpper, "BN"):
		prefix = "BN - "
	case strings.HasSuffix(skuUpper, "BX"):
		prefix = "BX - "
	case strings.HasSuffix(skuUpper, "C"):
		prefix = "Custom - "
	}

	// Product-line specific rules for SKUs ending in "B" mirror the current workbook behavior.
	if prefix == "" && strings.HasSuffix(skuUpper, "B") {
		pl := strings.TrimSpace(e.ProductLine)
		switch pl {
		case "2021":
			if strings.HasPrefix(strings.ToUpper(e.SKU), "FC") {
				prefix = "Bulk - "
			} else {
				prefix = "BX - "
			}
		case "BAS":
			prefix = "Bulk - "
		case "OAT":
			prefix = "BX - "
		}
	}

	if prefix != "" && !strings.HasPrefix(classDesc, prefix) {
		classDesc = prefix + classDesc
	}
	e.ClassDesc = classDesc
	return classDesc
}

// standardSheetCellFillColor calculates the current fill color for a standard-sheet cell based on
// MTO thresholds and the entry's status overrides.
func standardSheetCellFillColor(status string, columnIdx, mtoYtdIdx, mtoPyIdx int, mtoYTD, mtoPY float64, value interface{}) string {
	fillColor := "#FFFFFF"
	if (columnIdx == mtoYtdIdx || columnIdx == mtoPyIdx) && value != nil {
		if columnIdx == mtoYtdIdx {
			// MTO YTD uses lighter shades than MTO PY to keep the two columns visually distinct.
			if mtoYTD <= 1 {
				fillColor = "#FFCCCC"
			} else if mtoYTD <= 3 {
				fillColor = "#FFFFCC"
			} else {
				fillColor = "#CCFFCC"
			}
		} else {
			// MTO PY uses darker shades to match the historical-sales comparison column.
			if mtoPY <= 1 {
				fillColor = "#FF6666"
			} else if mtoPY <= 3 {
				fillColor = "#FFCC33"
			} else {
				fillColor = "#66FF66"
			}
		}
	}

	// Status-based shading still wins so rundown and discontinued items remain easy to spot.
	switch status {
	case "Rundown":
		fillColor = "#D3D3D3"
	case "Discontinued":
		fillColor = "#A9A9A9"
	}

	return fillColor
}

// applyStandardSheetWidths sets the column widths used by the standard report tabs.
func applyStandardSheetWidths(f *excelize.File, headers []string) error {
	for _, sheetName := range standardSheetNames {
		for i, h := range headers {
			col, _ := excelize.ColumnNumberToName(i + 1)
			if err := f.SetColWidth(sheetName, col, col, standardSheetWidthForHeader(h)); err != nil {
				return fmt.Errorf("failed to set width for %s column %s: %w", sheetName, col, err)
			}
		}
	}
	return nil
}

// standardSheetWidthForHeader returns the width used for one standard-sheet column header.
func standardSheetWidthForHeader(header string) float64 {
	switch header {
	case "Item Code":
		return 20
	case "QTY on Hand":
		return 12
	case "PO Num 1", "PO Num 2":
		return 12
	case "QTY on PO 1", "QTY on PO 2":
		return 12
	case "Total QTY on PO":
		return 15
	case "QTY on SO+BO":
		return 15
	case "QTY Available":
		return 15
	case "MTO YTD", "MTO PY":
		return 10
	case "QTY Sold+Issued YTD", "QTY Sold+Issued PY":
		return 20
	case "Class":
		return 20
	case "Status":
		return 15
	case "Occasion":
		return 20
	case "Description":
		return 35
	case "UPC":
		return 15
	case "Foil":
		return 10
	case "Royalty Code":
		return 15
	case "Dollar Sold YTD", "Dollar Sold PY":
		return 18
	default:
		return 12
	}
}

// applyStandardSheetFilters applies the autofilter range used by the standard report tabs.
func applyStandardSheetFilters(f *excelize.File, headers []string) error {
	if len(headers) == 0 {
		return fmt.Errorf("cannot apply autofilter to empty standard header set")
	}
	lastCol, _ := excelize.ColumnNumberToName(len(headers))
	for _, sheetName := range standardSheetNames {
		if err := f.AutoFilter(sheetName, fmt.Sprintf("A1:%s1", lastCol), nil); err != nil {
			return fmt.Errorf("failed to set autofilter for %s: %w", sheetName, err)
		}
	}
	return nil
}
