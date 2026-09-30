package hotsheet

import (
	"fmt"
	"sort"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	allProductsSheetName      = "All Products"
	ytdStockPrioritySheetName = "YTD Stock Priority"
)

// standardSheetNames lists the inventory tabs in workbook order.
var standardSheetNames = []string{allProductsSheetName, ytdStockPrioritySheetName}

// writeStandardSheets writes all inventory items to All Products and eligible items to
// YTD Stock Priority, sharing headers, formatting, widths, and filters. The PO flag controls
// detail columns; an error is returned if either worksheet cannot be written.
// YTD Stock Priority excludes rundown and discontinued items and sorts by MTO YTD.
func writeStandardSheets(f *excelize.File, entries []*inventoryEntry, hasPO bool) error {
	headers, mtoYtdIdx, mtoPyIdx := buildStandardSheetHeaders(hasPO)
	monthsThrough := currentMonthsThrough(time.Now())
	priorityEntries := make([]*inventoryEntry, 0, len(entries))
	for _, entry := range entries {
		if entry.Status != "Rundown" && entry.Status != "Discontinued" {
			priorityEntries = append(priorityEntries, entry)
		}
	}
	// Sorting a separate slice preserves the original inventory order on All Products.
	sort.SliceStable(priorityEntries, func(i, j int) bool {
		return standardMTOYTD(priorityEntries[i], monthsThrough) < standardMTOYTD(priorityEntries[j], monthsThrough)
	})
	for _, sheetName := range standardSheetNames {
		sheetEntries := entries
		if sheetName == ytdStockPrioritySheetName {
			sheetEntries = priorityEntries
		}
		if err := writeStandardSheetHeaders(f, sheetName, headers, hasPO); err != nil {
			return err
		}
		if err := writeStandardSheetRows(f, sheetName, sheetEntries, hasPO, monthsThrough, mtoYtdIdx, mtoPyIdx); err != nil {
			return err
		}
		if err := applyStandardSheetWidths(f, sheetName, headers); err != nil {
			return err
		}
		if err := applyStandardSheetFilters(f, sheetName, headers); err != nil {
			return err
		}
	}

	return nil
}

// buildStandardSheetHeaders returns the inventory header row, including Season before
// Occasion, Quantity Committed (sales orders plus back orders), and shipped
// YTD/PY quantities, plus MTO column indexes.
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
		"Quantity Committed",
		"QTY Available",
		"MTO YTD",
		"MTO PY",
		"QTY Shipped YTD",
		"QTY Shipped PY",
		"Class",
		"Status",
		"Season",
		"Occasion",
		"Description",
		"UPC",
		"Foil",
		"Card Size",
		"Royalty Code",
		"Dollar Sold YTD",
		"Dollar Sold PY",
		"Inactive",
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

// writeStandardSheetHeaders writes the supplied sheet's headers with the standard style
// and attaches explanatory MTO comments to their corresponding columns.
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
				Text:   "MTO YTD = QTY Available / ((QTY Shipped YTD + Quantity Committed) / monthsThrough + 1). Shipped = sold + nonnegative issued units. monthsThrough is the number of months completed in the current year (fractional). This shows months till out using year-to-date shipped pace including current sales orders/backorders.",
				Height: 190,
				Width:  200,
			}
			_ = f.AddComment(sheetName, cmt)
		}
		if h == "MTO PY" {
			cmt := excelize.Comment{
				Cell:   cell,
				Author: "Shane DuPrey",
				Text:   "MTO PY = QTY Available / (QTY Shipped PY / salesSeason + 1). Shipped = sold + nonnegative issued units. salesSeason used: Winter=6.5, Spring=5, Everyday=12. This shows months till out using prior-year shipped units scaled to the season length.",
				Height: 180,
				Width:  180,
			}
			_ = f.AddComment(sheetName, cmt)
		}
	}

	return nil
}

// writeStandardSheetRows writes the supplied inventory entries on one named worksheet,
// including their mapped Season before Occasion. Sales orders plus back orders are included
// in availability and YTD pace.
// The PO flag controls the columns, monthsThrough anchors MTO YTD, and the MTO column indexes
// select conditional coloring; an error is returned if a cell or style cannot be written.
// Season-specific MTO PY and class-prefix behavior are preserved.
func writeStandardSheetRows(f *excelize.File, sheetName string, entries []*inventoryEntry, hasPO bool, monthsThrough float64, mtoYtdIdx, mtoPyIdx int) error {
	rowIdx := 2
	for _, e := range entries {
		sh := mapOccasion(e.Occasion)

		// Determine the sales-season window used for MTO PY calculations.
		// Winter and Spring use their shorter merchandising seasons, while Everyday uses
		// the full year so the historical sales pace stays consistent with the workbook notes.
		var salesSeason float64
		switch sh {
		case "Winter":
			salesSeason = 6.5
		case "Spring":
			salesSeason = 5.0
		default:
			salesSeason = 12.0
		}

		// Inventory sold and issued quantities are separate source fields; combine
		// them once for displayed shipments and inventory-based MTO pace.
		committed := e.OnSO + e.OnBO
		totalInventory := e.OnHand + e.OnPO
		totalAvail := totalInventory - committed

		shippedYTD := e.YTDSold + max(e.YTDIssued, 0)
		shippedPY := e.SoldPY + max(e.IssuedPY, 0)
		shippedPerMonthPY := float64(shippedPY) / salesSeason

		mtoYTD := standardMTOYTD(e, monthsThrough)
		mtoPY := float64(totalAvail) / (shippedPerMonthPY + 1)

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
			committed,
			totalAvail,
			mtoYTD,
			mtoPY,
			shippedYTD,
			shippedPY,
			classDesc,
			e.Status,
			sh,
			e.Occasion,
			e.Description,
			e.UPC,
			e.Foil,
			e.CardSize,
			e.RoyaltyCode,
			e.DollarSoldYTD,
			e.DollarSoldPY,
			e.Inactive,
		)

		dollarYTDCol := len(vals) - 3
		dollarPYCol := len(vals) - 2

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

// standardMTOYTD returns an entry's months-to-out value from available stock and the
// year-to-date shipped pace including committed orders, using the supplied monthsThrough.
func standardMTOYTD(e *inventoryEntry, monthsThrough float64) float64 {
	committed := e.OnSO + e.OnBO
	available := e.OnHand + e.OnPO - committed
	shippedYTD := e.YTDSold + max(e.YTDIssued, 0)
	return float64(available) / ((float64(shippedYTD)+float64(committed))/monthsThrough + 1)
}

// applyStandardDisplayClassPrefix applies the display-time class prefix rules while keeping
// RawClassDesc unchanged. Repeated rendering of an entry must not add the prefix twice.
func applyStandardDisplayClassPrefix(e *inventoryEntry) string {
	classDesc := strings.TrimSpace(e.ClassDesc)
	skuUpper := strings.ToUpper(strings.TrimSpace(e.SKU))
	prefix := ""

	// Check longer or more specific suffixes first so the display prefix is deterministic.
	switch {
	case strings.HasSuffix(skuUpper, "-LLB") || strings.HasSuffix(skuUpper, "LLB"):
		prefix = "LLB - "
	case strings.HasSuffix(skuUpper, "-TB") || strings.HasSuffix(skuUpper, "TB") || strings.HasPrefix(skuUpper, "TB") || strings.HasSuffix(skuUpper, "TBB"):
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

	// An empty class can leave just the prefix after the first rendering; retain
	// its trailing space so both worksheets show the same displayed class.
	if prefix != "" && classDesc == strings.TrimSpace(prefix) {
		classDesc = prefix
	} else if prefix != "" && !strings.HasPrefix(classDesc, prefix) {
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

// applyStandardSheetWidths sets the column widths for one sheet's headers, including any
// extra combined-sheet columns; it returns a width-setting error if Excelize fails.
func applyStandardSheetWidths(f *excelize.File, sheetName string, headers []string) error {
	for i, h := range headers {
		col, _ := excelize.ColumnNumberToName(i + 1)
		if err := f.SetColWidth(sheetName, col, col, standardSheetWidthForHeader(h)); err != nil {
			return fmt.Errorf("failed to set width for %s column %s: %w", sheetName, col, err)
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
	case "Quantity Committed":
		return 22
	case "QTY Available":
		return 15
	case "MTO YTD", "MTO PY":
		return 10
	case "QTY Shipped YTD", "QTY Shipped PY":
		return 20
	case "Class":
		return 20
	case "Status", "Season":
		return 15
	case "Occasion":
		return 20
	case "Description":
		return 35
	case "UPC":
		return 15
	case "Foil":
		return 10
	case "Card Size":
		return 10
	case "Royalty Code":
		return 15
	case "Dollar Sold YTD", "Dollar Sold PY":
		return 18
	default:
		return 12
	}
}

// applyStandardSheetFilters filters all columns for the given sheet's headers, returning
// an error if the header list is empty or Excelize cannot set the filter.
func applyStandardSheetFilters(f *excelize.File, sheetName string, headers []string) error {
	if len(headers) == 0 {
		return fmt.Errorf("cannot apply autofilter to empty standard header set")
	}
	lastCol, _ := excelize.ColumnNumberToName(len(headers))
	if err := f.AutoFilter(sheetName, fmt.Sprintf("A1:%s1", lastCol), nil); err != nil {
		return fmt.Errorf("failed to set autofilter for %s: %w", sheetName, err)
	}
	return nil
}
