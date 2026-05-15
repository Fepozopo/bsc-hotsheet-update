package hotsheet

import (
	"fmt"
	"log/slog"
	"path/filepath"
	"sort"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Shared workbook column lookups keep the row-parsing helpers aligned with the source report layout.
var (
	inventorySKUIdx         = colToIndex("B")
	inventoryProductLineIdx = colToIndex("B")
	inventoryClassIdx       = colToIndex("D")
	inventoryStatusIdx      = colToIndex("F")
	inventoryOnHandIdx      = colToIndex("H")
	inventoryOnPOIdx        = colToIndex("J")
	inventoryOnSOIdx        = colToIndex("L")
	inventoryOnBOIdx        = colToIndex("N")
	inventoryTotalAvailIdx  = colToIndex("P")
	inventoryYTDSoldIdx     = colToIndex("R")
	inventoryYTDIssuedIdx   = colToIndex("T")
	inventorySoldPYIdx      = colToIndex("V")
	inventoryIssuedPYIdx    = colToIndex("X")
	inventoryFoilIdx        = colToIndex("Z")
	inventoryOccasionIdx    = colToIndex("AB")
	inventoryDescIdx        = colToIndex("AD")
	inventoryUPCIdx         = colToIndex("AF")
	inventoryRoyaltyCodeIdx = colToIndex("AH")
	inventoryDollarYTDIdx   = colToIndex("AJ")
	inventoryDollarPYIdx    = colToIndex("AL")
	poDataIdx               = colToIndex("A")
	poStatusIdx             = colToIndex("G")
	poOnPOIdx               = colToIndex("I")
	poOnPOBackorderIdx      = colToIndex("K")
)

// standardSheetNames preserves the original workbook tab order used by the hotsheet.
var standardSheetNames = []string{"Everyday", "Winter", "Spring"}

// loadInventoryEntries opens the inventory workbook, parses the inventory rows, and returns
// the populated inventory map keyed by SKU.
func loadInventoryEntries(inventoryPath string, logger *slog.Logger) (map[string]*entry, error) {
	if logger != nil {
		logger.Info("loading inventory report", "path", inventoryPath)
	}

	wbInv, err := excelize.OpenFile(inventoryPath)
	if err != nil {
		return nil, fmt.Errorf("failed to open inventory report %s: %w", inventoryPath, err)
	}
	defer func() {
		_ = wbInv.Close()
	}()

	invRows, err := wbInv.GetRows("Sheet1")
	if err != nil {
		return nil, fmt.Errorf("failed to read inventory sheet: %w", err)
	}
	if len(invRows) < 1 {
		return nil, fmt.Errorf("inventory report appears empty")
	}

	invMap := make(map[string]*entry)
	for rowNum := 2; ; rowNum += 3 {
		e, stop := parseInventoryEntry(invRows, rowNum, logger)
		if stop {
			break
		}
		if e == nil {
			continue
		}
		invMap[e.SKU] = e
	}

	return invMap, nil
}

// parseInventoryEntry parses one inventory item from the worksheet rows and reports whether
// parsing should stop because the scan reached the workbook footer or ran out of rows.
func parseInventoryEntry(rows [][]string, rowNum int, logger *slog.Logger) (*entry, bool) {
	if rowNum-1 >= len(rows) {
		return nil, true
	}

	// Item codes start at B2 and repeat every 3 rows in the inventory export.
	sku := getCellAt(rows, rowNum, inventorySKUIdx)
	if sku == "" {
		if logger != nil {
			logger.Info("Skipping empty SKU at inventory row", "row", rowNum)
		}
		return nil, false
	}
	// If the SKU cell looks like a run date, stop parsing so footer metadata is ignored.
	if isRunDate(sku) {
		if logger != nil {
			logger.Info("Encountered run-date/footer, stopping parse", "value", sku, "row", rowNum)
		}
		return nil, true
	}

	valRow := rowNum + 2
	if valRow-1 >= len(rows) {
		return nil, true
	}

	e := &entry{SKU: sku}
	e.ProductLine = getCellAt(rows, valRow, inventoryProductLineIdx)
	e.ClassDesc = getCellAt(rows, valRow, inventoryClassIdx)
	e.RawClassDesc = e.ClassDesc
	e.Status = getCellAt(rows, valRow, inventoryStatusIdx)
	e.OnHand = parseInt(getCellAt(rows, valRow, inventoryOnHandIdx))
	e.OnPO = parseInt(getCellAt(rows, valRow, inventoryOnPOIdx))
	e.OnSO = parseInt(getCellAt(rows, valRow, inventoryOnSOIdx))
	e.OnBO = parseInt(getCellAt(rows, valRow, inventoryOnBOIdx))
	e.TotalAvailable = parseInt(getCellAt(rows, valRow, inventoryTotalAvailIdx))
	e.YTDSold = parseInt(getCellAt(rows, valRow, inventoryYTDSoldIdx))
	e.YTDIssued = parseInt(getCellAt(rows, valRow, inventoryYTDIssuedIdx))
	e.SoldPY = parseInt(getCellAt(rows, valRow, inventorySoldPYIdx))
	e.IssuedPY = parseInt(getCellAt(rows, valRow, inventoryIssuedPYIdx))
	e.Foil = getCellAt(rows, valRow, inventoryFoilIdx)
	e.Occasion = getCellAt(rows, valRow, inventoryOccasionIdx)
	e.Description = getCellAt(rows, valRow, inventoryDescIdx)
	e.UPC = getCellAt(rows, valRow, inventoryUPCIdx)
	e.RoyaltyCode = getCellAt(rows, valRow, inventoryRoyaltyCodeIdx)
	e.DollarSoldYTD = parseInventoryDollar(getCellAt(rows, valRow, inventoryDollarYTDIdx))
	e.DollarSoldPY = parseInventoryDollar(getCellAt(rows, valRow, inventoryDollarPYIdx))

	if logger != nil {
		logger.Debug("Inventory parse",
			"SKU", e.SKU,
			"skuRow", rowNum,
			"valRow", valRow,
			"ProductLine", e.ProductLine,
			"ClassDesc", e.ClassDesc,
			"Status", e.Status,
			"OnHand", e.OnHand,
			"OnPO", e.OnPO,
			"OnSO", e.OnSO,
			"OnBO", e.OnBO,
			"TotalAvailable", e.TotalAvailable,
			"YTDSold", e.YTDSold,
			"YTDIssued", e.YTDIssued,
			"SoldPY", e.SoldPY,
			"IssuedPY", e.IssuedPY,
			"Foil", e.Foil,
			"Occasion", e.Occasion,
			"Description", e.Description,
			"UPC", e.UPC,
			"RoyaltyCode", e.RoyaltyCode,
			"DollarSoldYTD", e.DollarSoldYTD,
			"DollarSoldPY", e.DollarSoldPY,
		)
	}

	return e, false
}

// parseInventoryDollar converts inventory currency text into a float64 while preserving the
// forgiving parsing used by the current workbook import.
func parseInventoryDollar(s string) float64 {
	trimmed := strings.TrimSpace(s)
	if trimmed == "" {
		return 0
	}
	trimmed = strings.ReplaceAll(trimmed, "$", "")
	trimmed = strings.ReplaceAll(trimmed, ",", "")
	trimmed = strings.ReplaceAll(trimmed, "(", "-")
	trimmed = strings.ReplaceAll(trimmed, ")", "")
	return parseFloat(trimmed)
}

// mergePOData opens the optional PO workbook and merges PO numbers and quantities into the
// provided inventory map. PO-only SKUs are skipped so the workbook does not create UNKNOWN groups.
func mergePOData(poPath string, invMap map[string]*entry, logger *slog.Logger) error {
	if strings.TrimSpace(poPath) == "" {
		return nil
	}

	wbPO, err := excelize.OpenFile(poPath)
	if err != nil {
		return fmt.Errorf("failed to open PO report %s: %w", poPath, err)
	}
	defer func() {
		_ = wbPO.Close()
	}()

	poRows, err := wbPO.GetRows("Sheet1")
	if err != nil {
		return fmt.Errorf("failed to read PO sheet: %w", err)
	}
	if len(poRows) == 0 {
		return nil
	}

	for rowNum := 1; rowNum <= len(poRows); rowNum++ {
		row := poRows[rowNum-1]
		if poDataIdx >= len(row) {
			continue
		}

		sku := strings.TrimSpace(row[poDataIdx])
		if sku == "" {
			continue
		}

		e, ok := invMap[sku]
		if !ok {
			if logger != nil {
				logger.Info("Skipping PO-only SKU (not present in inventory)", "SKU", sku)
			}
			continue
		}

		// Walk subsequent rows until we hit a line that starts with "Item" (end of section)
		// or we've collected up to two PO lines for this SKU.
		maxPOs := 2
		poCount := 0
		for scanRow := rowNum + 1; scanRow <= len(poRows) && poCount < maxPOs; scanRow++ {
			nextRow := getRow(poRows, scanRow)
			if nextRow == nil {
				break
			}

			poCell := strings.TrimSpace(getCell(nextRow, poDataIdx))
			if poCell == "" {
				// skip empty lines
				continue
			}
			if strings.HasPrefix(strings.ToUpper(poCell), "ITEM") {
				// end of PO block for this SKU
				break
			}

			poNum, qty := applyPOToEntry(e, nextRow)
			if logger != nil {
				status := strings.TrimSpace(getCell(nextRow, poStatusIdx))
				logger.Debug("Individual PO parse",
					"SKU", sku,
					"skuRow", rowNum,
					"poRow", scanRow,
					"PO Num", poNum,
					"QTY", qty,
					"PO Status", status,
				)
			}

			poCount++
		}
	}

	return nil
}

// applyPOToEntry normalizes one PO line, chooses the correct quantity column based on the status,
// and assigns the PO into the first available slot on the inventory entry.
func applyPOToEntry(e *entry, poRow []string) (string, int) {
	status := strings.TrimSpace(getCell(poRow, poStatusIdx))
	var qty int
	if strings.EqualFold(status, "Back Order") {
		qty = parseInt(getCell(poRow, poOnPOBackorderIdx))
	} else {
		qty = parseInt(getCell(poRow, poOnPOIdx))
	}

	// Normalize PO numbers by removing leading zeros while preserving a usable zero value.
	poNum := strings.TrimLeft(strings.TrimSpace(getCell(poRow, poDataIdx)), "0")
	if poNum == "" {
		poNum = "0"
	}
	assignPO(e, poNum, qty)
	return poNum, qty
}

// groupEntriesByProductLine iterates the inventory map and groups entries by ProductLine while
// preserving the current behavior of skipping blank product-line entries.
func groupEntriesByProductLine(invMap map[string]*entry, logger *slog.Logger) map[string][]*entry {
	productGroups := make(map[string][]*entry)
	for _, e := range invMap {
		if e == nil {
			continue
		}
		pl := strings.TrimSpace(e.ProductLine)
		if pl == "" {
			if logger != nil {
				// Skip entries without a ProductLine so the workbook does not create UNKNOWN files.
				logger.Info("Skipping SKU with empty ProductLine (likely PO-only entry)", "SKU", e.SKU)
			}
			continue
		}
		productGroups[pl] = append(productGroups[pl], e)
	}
	return productGroups
}

// sortEntriesForProductLine sorts a product-line slice into a stable SKU order before workbook
// generation so the tab contents remain consistent.
func sortEntriesForProductLine(entries []*entry) {
	sort.SliceStable(entries, func(i, j int) bool { return entries[i].SKU < entries[j].SKU })
}

// buildProductLineWorkbook creates one workbook for a product line, writes the standard report
// sheets and Data Insights sheet, and saves the result to disk.
func buildProductLineWorkbook(productLine string, entries []*entry, outputDir, dateStr string, hasPO bool, logger *slog.Logger) (string, error) {
	f := newProductLineWorkbook()
	defer func() {
		_ = f.Close()
	}()

	if err := writeStandardSheets(f, entries, hasPO); err != nil {
		if logger != nil {
			logger.Error("failed to write standard sheets", "productLine", productLine, "err", err)
		}
		return "", fmt.Errorf("failed to write standard sheets for %s: %w", productLine, err)
	}

	if err := writeDataInsightsSheet(f, entries); err != nil {
		if logger != nil {
			logger.Error("failed to create Data Insights sheet", "productLine", productLine, "err", err)
		}
		return "", fmt.Errorf("failed to create Data Insights sheet for %s: %w", productLine, err)
	}

	outPath, err := saveWorkbook(f, outputDir, productLine, dateStr)
	if err != nil {
		if logger != nil {
			logger.Error("failed to save hotsheet for product line", "productLine", productLine, "err", err)
		}
		return "", err
	}

	return outPath, nil
}

// newProductLineWorkbook creates the workbook shell used for each product-line export.
func newProductLineWorkbook() *excelize.File {
	f := excelize.NewFile()
	idx, _ := f.NewSheet("Everyday")
	f.SetActiveSheet(idx)
	_, _ = f.NewSheet("Winter")
	_, _ = f.NewSheet("Spring")
	// Delete the default Sheet1 if it still exists so the output matches the existing workbook layout.
	if idxSheet, _ := f.GetSheetIndex("Sheet1"); idxSheet != -1 {
		_ = f.DeleteSheet("Sheet1")
	}
	return f
}

// writeStandardSheets writes the Everyday, Winter, and Spring tabs, their headers, their rows,
// and the shared widths and filters used by the standard hotsheet layout.
func writeStandardSheets(f *excelize.File, entries []*entry, hasPO bool) error {
	headers, mtoYtdIdx, mtoPyIdx := buildStandardSheetHeaders(hasPO)

	for _, sheetName := range standardSheetNames {
		if err := writeStandardSheetHeaders(f, sheetName, headers, hasPO); err != nil {
			return err
		}
	}

	monthsThrough := currentMonthsThrough()
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
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#E6E6FA"}, Pattern: 1},
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
				Text:   "MTO YTD + QTY Available / ((QTY Sold+Issued YTD) / (monthsThrough + 1)). monthsThrough is the number of months completed in the current year (fractional). This shows months till out using year-to-date sales pace.",
				Height: 190,
				Width:  200,
			}
			_ = f.AddComment(sheetName, cmt)
		}
		if h == "MTO PY" {
			cmt := excelize.Comment{
				Cell:   cell,
				Author: "Shane DuPrey",
				Text:   "MTO PY = QTY Available / ((QTY Sold+Issued PY) / (salesSeason + 1)). salesSeason used: Winter=6, Spring=5, Everyday=12. This shows months till out using prior-year sales scaled to the season length.",
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
func writeStandardSheetRows(f *excelize.File, sheetName string, entries []*entry, hasPO bool, monthsThrough float64, mtoYtdIdx, mtoPyIdx int) error {
	rowIdx := 2
	for _, e := range entries {
		sh := mapOccasion(e.Occasion)
		if sh != sheetName {
			continue
		}

		// Determine the sales-season window used for MTO PY calculations.
		salesSeason := 12.0
		switch sh {
		case "Winter":
			salesSeason = 6.0
		case "Spring":
			salesSeason = 5.0
		default:
			salesSeason = 12.0
		}

		// Preserve the current derived-value math exactly so workbook output stays compatible.
		onSOBO := e.OnSO + e.OnBO
		totalAvail := e.OnHand + e.OnPO - onSOBO
		totalSoldYTD := e.YTDSold + max(e.YTDIssued, 0)
		totalSoldPY := e.SoldPY + max(e.IssuedPY, 0)
		soldPerMonthYTD := float64(totalSoldYTD) / monthsThrough
		soldPerMonthPY := float64(totalSoldPY) / salesSeason
		mtoYTD := float64(totalAvail) / (soldPerMonthYTD + 1)
		mtoPY := float64(totalAvail) / (soldPerMonthPY + 1)

		classDesc := applyStandardClassPrefix(e)

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

		for c, v := range vals {
			cell, _ := excelize.CoordinatesToCellName(c+1, rowIdx)
			if err := f.SetCellValue(sheetName, cell, v); err != nil {
				return fmt.Errorf("failed to write %s cell %s: %w", sheetName, cell, err)
			}

			fillColor := standardSheetCellFillColor(e.Status, c, mtoYtdIdx, mtoPyIdx, mtoYTD, mtoPY, v)
			style, err := f.NewStyle(&excelize.Style{
				Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
				Border: []excelize.Border{
					{Type: "left", Color: "000000", Style: 1},
					{Type: "right", Color: "000000", Style: 1},
					{Type: "top", Color: "000000", Style: 1},
					{Type: "bottom", Color: "000000", Style: 1},
				},
				Fill: excelize.Fill{Type: "pattern", Color: []string{fillColor}, Pattern: 1},
			})
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

// applyStandardClassPrefix applies the current display-time class prefix rules while keeping the
// original inventory class available through RawClassDesc for downstream reuse.
func applyStandardClassPrefix(e *entry) string {
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

// saveWorkbook builds the final output path, sanitizes the product line name, and writes the
// workbook to disk.
func saveWorkbook(f *excelize.File, outputDir, productLine, dateStr string) (string, error) {
	outDir := outputDir
	if strings.TrimSpace(outDir) == "" {
		outDir = "."
	}
	fileName := fmt.Sprintf("%s_hotsheet_%s.xlsx", sanitizeFileName(productLine), dateStr)
	outPath := filepath.Join(outDir, fileName)
	if err := f.SaveAs(outPath); err != nil {
		return "", fmt.Errorf("failed to save hotsheet %s: %w", outPath, err)
	}
	return outPath, nil
}
