package hotsheet

import (
	"fmt"
	"path/filepath"
	"sort"
	"strings"
	"time"

	helpers "github.com/Fepozopo/bsc-hotsheet-update/helpers"
	"github.com/xuri/excelize/v2"
)

// CreateFromReports parses the inventory and PO reports and generates one hotsheet Excel
// file per Product Line. Files are written to outputDir (or current directory if empty)
// and follow the naming pattern: {Product Line}_hotsheet_YYYYMMDD.xlsx
// Returns the list of generated file paths.
func CreateFromReports(inventoryPath, poPath, outputDir string) ([]string, error) {
	// logger for the operation
	logger, logFile, err := helpers.CreateLogger("create", "all", "", "INFO")
	if err != nil {
		return nil, fmt.Errorf("failed to create logger: %w", err)
	}
	defer logFile.Close()

	// Open inventory workbook
	wbInv, err := excelize.OpenFile(inventoryPath)
	if err != nil {
		return nil, fmt.Errorf("failed to open inventory report %s: %w", inventoryPath, err)
	}
	defer wbInv.Close()
	invSheet := "Sheet1"
	invRows, err := wbInv.GetRows(invSheet)
	if err != nil {
		return nil, fmt.Errorf("failed to read inventory sheet: %w", err)
	}
	if len(invRows) < 1 {
		return nil, fmt.Errorf("inventory report appears empty")
	}

	invMap := make(map[string]*entry)

	// Inventory report has no headers. Item codes start at B2 and repeat every 3 rows.
	// Column-letter mapping for value row (2 rows below the SKU row)
	skuCol := "B"
	plCol := "B"
	classCol := "D"
	statusCol := "F"
	onHandCol := "H"
	onPOCol := "J"
	onSOCol := "L"
	onBOCol := "N"
	totalAvailCol := "P"
	ytdSoldCol := "R"
	ytdIssuedCol := "T"
	soldPYCol := "V"
	issuedPYCol := "X"
	foilCol := "Z"
	occasionCol := "AB"
	descCol := "AD"
	upcCol := "AF"

	skuIdx := colToIndex(skuCol)
	plIdx := colToIndex(plCol)
	classIdx := colToIndex(classCol)
	statusIdx := colToIndex(statusCol)
	onHandIdx := colToIndex(onHandCol)
	onPOIdx := colToIndex(onPOCol)
	onSOIdx := colToIndex(onSOCol)
	onBOIdx := colToIndex(onBOCol)
	totalAvailIdx := colToIndex(totalAvailCol)
	ytdSoldIdx := colToIndex(ytdSoldCol)
	ytdIssuedIdx := colToIndex(ytdIssuedCol)
	soldPYIdx := colToIndex(soldPYCol)
	issuedPYIdx := colToIndex(issuedPYCol)
	foilIdx := colToIndex(foilCol)
	occasionIdx := colToIndex(occasionCol)
	descIdx := colToIndex(descCol)
	upcIdx := colToIndex(upcCol)

	startRow := 2
	for r := startRow; ; r += 3 {
		// SKU row is at column B, row r
		if r-1 >= len(invRows) {
			break
		}
		sku := getCellAt(invRows, r, skuIdx)
		// skip empty SKU rows and continue scanning; do not break here because blank rows may appear
		if sku == "" {
			logger.Printf("Skipping empty SKU at inventory row %d", r)
			continue
		}
		// if the SKU cell looks like a run date, stop parsing
		if isRunDate(sku) {
			logger.Printf("Encountered run-date/footer '%s' at inventory row %d — stopping parse", sku, r)
			break
		}

		valRow := r + 2
		if valRow-1 >= len(invRows) {
			break
		}

		e := &entry{SKU: sku}
		e.ProductLine = getCellAt(invRows, valRow, plIdx)
		e.ClassDesc = getCellAt(invRows, valRow, classIdx)
		e.Status = getCellAt(invRows, valRow, statusIdx)
		e.OnHand = parseInt(getCellAt(invRows, valRow, onHandIdx))
		e.OnPO = parseInt(getCellAt(invRows, valRow, onPOIdx))
		e.OnSO = parseInt(getCellAt(invRows, valRow, onSOIdx))
		e.OnBO = parseInt(getCellAt(invRows, valRow, onBOIdx))
		e.TotalAvailable = parseInt(getCellAt(invRows, valRow, totalAvailIdx))
		e.YTDSold = parseInt(getCellAt(invRows, valRow, ytdSoldIdx))
		e.YTDIssued = parseInt(getCellAt(invRows, valRow, ytdIssuedIdx))
		e.SoldPY = parseInt(getCellAt(invRows, valRow, soldPYIdx))
		e.IssuedPY = parseInt(getCellAt(invRows, valRow, issuedPYIdx))
		e.Foil = getCellAt(invRows, valRow, foilIdx)
		e.Occasion = getCellAt(invRows, valRow, occasionIdx)
		e.Description = getCellAt(invRows, valRow, descIdx)
		e.UPC = getCellAt(invRows, valRow, upcIdx)

		logger.Printf("Inventory parse: SKU=%s skuRow=%d valRow=%d ProductLine=%s OnHand=%d OnPO=%d", e.SKU, r, valRow, e.ProductLine, e.OnHand, e.OnPO)

		invMap[e.SKU] = e
	}

	// Parse PO report to capture PO numbers and per-PO quantities (do not modify inventory OnPO)
	if poPath != "" {
		wbPO, err := excelize.OpenFile(poPath)
		if err != nil {
			logger.Printf("failed to open PO report: %v", err)
		} else {
			defer wbPO.Close()
			poRows, err := wbPO.GetRows("Sheet1")
			if err != nil {
				logger.Printf("failed to read PO sheet: %v", err)
			} else if len(poRows) > 0 {
				dataCol := "A"
				onPOCol := "I"
				onPOBackorderCol := "K"
				poStatusCol := "G"

				dataIdx := colToIndex(dataCol)
				onPOIdx := colToIndex(onPOCol)
				onPOBackIdx := colToIndex(onPOBackorderCol)
				poStatusIdx := colToIndex(poStatusCol)

				for rowNum := 1; rowNum < len(poRows)+1; rowNum++ {
					row := poRows[rowNum-1]
					if dataIdx >= len(row) {
						continue
					}
					sku := strings.TrimSpace(row[dataIdx])
					if sku == "" {
						continue
					}

					// ensure inventory entry exists; skip PO-only SKUs (avoid creating UNKNOWN product-line groups)
					e, ok := invMap[sku]
					if !ok {
						logger.Printf("Skipping PO-only SKU %s (not present in inventory)", sku)
						continue
					}

					// Walk subsequent rows until we hit a line that starts with "Item" (end of section)
					// or we've collected up to three PO lines for this SKU.
					maxPOs := 3
					poCount := 0
					for r := rowNum + 1; r <= len(poRows) && poCount < maxPOs; r++ {
						nextRow := getRow(poRows, r)
						if nextRow == nil {
							break
						}
						poCell := strings.TrimSpace(getCell(nextRow, dataIdx))
						if poCell == "" {
							// skip empty lines
							continue
						}
						if strings.HasPrefix(strings.ToUpper(poCell), "ITEM") {
							// end of PO block for this SKU
							break
						}
						status := strings.TrimSpace(getCell(nextRow, poStatusIdx))
						var qty int
						if strings.EqualFold(status, "Back Order") {
							qty = parseInt(getCell(nextRow, onPOBackIdx))
						} else {
							qty = parseInt(getCell(nextRow, onPOIdx))
						}
						assignPO(e, poCell, qty)
						poCount++
					}
				}
			}
		}
	}

	// group entries by product line
	productGroups := make(map[string][]*entry)
	for _, e := range invMap {
		pl := strings.TrimSpace(e.ProductLine)
		if pl == "" {
			// skip entries without a ProductLine (likely discovered only from PO); this avoids many UNKNOWN files
			logger.Printf("Skipping SKU %s with empty ProductLine (likely PO-only entry)", e.SKU)
			continue
		}
		productGroups[pl] = append(productGroups[pl], e)
	}

	headersOut := []string{"Item Code", "Product Line", "Class Description", "Status", "Quantity on Hand", "Quantity on Purchase Order", "PO Number 1", "Quantity on PO 1", "PO Number 2", "Quantity on PO 2", "PO Number 3", "Quantity on PO 3", "Quantity on Sales Order", "Quantity on Back Order", "Total Quantity Available", "Quantity Sold YTD", "Quantity Issued YTD", "Quantity Sold PY", "Quantity Issued PY", "Foil", "Occasion", "Description", "UPC"}

	var outputs []string
	dateStr := time.Now().Format("20060102")

	for pl, entries := range productGroups {
		// sort by SKU
		sort.Slice(entries, func(i, j int) bool { return entries[i].SKU < entries[j].SKU })

		f := excelize.NewFile()
		// create sheets: Everyday, Winter, Spring
		idx, _ := f.NewSheet("Everyday")
		f.SetActiveSheet(idx)
		_, _ = f.NewSheet("Winter")
		_, _ = f.NewSheet("Spring")
		// delete default Sheet1 if still present
		if idxSheet, _ := f.GetSheetIndex("Sheet1"); idxSheet != 0 {
			_ = f.DeleteSheet("Sheet1")
		}

		// write headers to each sheet
		for _, sh := range []string{"Everyday", "Winter", "Spring"} {
			for c, h := range headersOut {
				cell, _ := excelize.CoordinatesToCellName(c+1, 1)
				f.SetCellValue(sh, cell, h)
			}
		}

		// row counters per sheet
		rowIdx := map[string]int{"Everyday": 2, "Winter": 2, "Spring": 2}

		for _, e := range entries {
			// determine sheet
			sh := mapOccasion(e.Occasion)
			// compute derived fields
			onSOBO := e.OnSO + e.OnBO
			totalAvail := e.OnHand + e.OnPO - onSOBO
			// write row (include per-PO details)
			vals := []interface{}{e.SKU, pl, e.ClassDesc, e.Status, e.OnHand, e.OnPO, e.PONum1, e.OnPO1, e.PONum2, e.OnPO2, e.PONum3, e.OnPO3, e.OnSO, e.OnBO, totalAvail, e.YTDSold, e.YTDIssued, e.SoldPY, e.IssuedPY, e.Foil, e.Occasion, e.Description, e.UPC}
			r := rowIdx[sh]
			for c, v := range vals {
				cell, _ := excelize.CoordinatesToCellName(c+1, r)
				f.SetCellValue(sh, cell, v)
			}
			rowIdx[sh] = r + 1
		}

		// ensure output directory
		outDir := outputDir
		if outDir == "" {
			outDir = "."
		}
		fileName := fmt.Sprintf("%s_hotsheet_%s.xlsx", sanitizeFileName(pl), dateStr)
		outPath := filepath.Join(outDir, fileName)
		if err := f.SaveAs(outPath); err != nil {
			logger.Printf("failed to save hotsheet for %s: %v", pl, err)
			return outputs, fmt.Errorf("failed to save hotsheet %s: %w", outPath, err)
		}
		outputs = append(outputs, outPath)
	}

	return outputs, nil
}
