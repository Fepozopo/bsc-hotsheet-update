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
					// or we've collected up to two PO lines for this SKU.
					maxPOs := 2
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
						// normalize PO number by removing leading zeros
						poNum := strings.TrimLeft(strings.TrimSpace(poCell), "0")
						if poNum == "" {
							// if result is empty (e.g., original was "0000"), set to "0"
							poNum = "0"
						}
						assignPO(e, poNum, qty)
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

	// Build headers. If there's no PO report provided, omit per-PO and SO/BO columns.
	hasPO := poPath != ""
	headersOut := []string{
		"Item Code",
		"QTY on Hand",
		"Total QTY on PO",
	}
	if hasPO {
		headersOut = append(headersOut,
			"PO Num 1",
			"QTY on PO 1",
			"PO Num 2",
			"QTY on PO 2",
			"QTY on SO/BO",
		)
	}
	headersOut = append(headersOut,
		"QTY Available",
		"MTO YTD",
		"MTO PY",
		"QTY Sold/Issued YTD",
		"QTY Sold/Issued PY",
		"Class",
		"Status",
		"Occasion",
		"Description",
		"UPC",
		"Foil",
	)
	// determine the index positions of MTO columns so coloring logic below works regardless
	mtoYtdIdx, mtoPyIdx := -1, -1
	for i, h := range headersOut {
		if h == "MTO YTD" {
			mtoYtdIdx = i
		}
		if h == "MTO PY" {
			mtoPyIdx = i
		}
	}

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
		if idxSheet, _ := f.GetSheetIndex("Sheet1"); idxSheet != -1 {
			_ = f.DeleteSheet("Sheet1")
		}

		// write headers to each sheet
		for _, sh := range []string{"Everyday", "Winter", "Spring"} {
			for c, h := range headersOut {
				cell, _ := excelize.CoordinatesToCellName(c+1, 1)
				f.SetCellValue(sh, cell, h)
				// make header bold, centered, with borders
				style, _ := f.NewStyle(&excelize.Style{
					Font:      &excelize.Font{Bold: true},
					Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
					Border: []excelize.Border{
						{Type: "left", Color: "000000", Style: 1},
						{Type: "right", Color: "000000", Style: 1},
						{Type: "top", Color: "000000", Style: 1},
						{Type: "bottom", Color: "000000", Style: 1},
					},
				})
				f.SetCellStyle(sh, cell, cell, style)
			}
		}

		// row counters per sheet
		rowIdx := map[string]int{"Everyday": 2, "Winter": 2, "Spring": 2}

		// compute months through the year for MTO calculation; use current date to determine fraction of month completed
		now := time.Now()
		year := now.Year()
		month := now.Month()
		// number of days in the current month
		daysInMonth := time.Date(year, month+1, 0, 0, 0, 0, 0, now.Location()).Day()
		// months completed plus fraction of current month (e.g., June 15 -> 5 + 15/30 = 5.5)
		monthsThrough := float64(int(month)-1) + float64(now.Day())/float64(daysInMonth)
		// guard against zero (shouldn't happen), fallback to 1 month
		if monthsThrough <= 0 {
			monthsThrough = 1
		}

		for _, e := range entries {
			// determine sheet
			sh := mapOccasion(e.Occasion)
			// determine sales season
			salesSeason := 12.0 // default to full year
			switch sh {
			case "Winter":
				salesSeason = 6.0 // typically most sales occur in last 6 months of year
			case "Spring":
				salesSeason = 4.0 // typically most sales occur in first 4 months of year
			default:
				salesSeason = 12.0 // assume sales spread evenly across the year for Everyday
			}

			// compute derived fields
			onSOBO := e.OnSO + e.OnBO
			totalAvail := e.OnHand + e.OnPO - onSOBO
			soldIssuedYTD := e.YTDSold + max(e.YTDIssued, 0)
			soldIssuedPY := e.SoldPY + max(e.IssuedPY, 0)

			// MTO = Months Till Out
			siPerMonthYTD := float64(soldIssuedYTD) / monthsThrough
			siPerMonthPY := float64(soldIssuedPY) / salesSeason
			mtoYTD := float64(totalAvail) / (siPerMonthYTD + 1)
			mtoPY := float64(totalAvail) / (siPerMonthPY + 1)

			// write row (include per-PO details)
			// Build row values dynamically to match headersOut (PO columns omitted if no PO report)
			vals := []interface{}{
				e.SKU,
				e.OnHand,
				e.OnPO,
			}
			if hasPO {
				vals = append(vals, e.PONum1, e.OnPO1, e.PONum2, e.OnPO2, onSOBO)
			}
			vals = append(vals,
				totalAvail,
				mtoYTD,
				mtoPY,
				soldIssuedYTD,
				soldIssuedPY,
				e.ClassDesc,
				e.Status,
				e.Occasion,
				e.Description,
				e.UPC,
				e.Foil,
			)
			r := rowIdx[sh]
			for c, v := range vals {
				cell, _ := excelize.CoordinatesToCellName(c+1, r)
				f.SetCellValue(sh, cell, v)

				fillColor := "#FFFFFF" // default white
				// fill MTO columns based on thresholds: red if <=1 month, yellow if <=3 months, otherwise white
				if (c == mtoYtdIdx || c == mtoPyIdx) && v != nil {
					if c == mtoYtdIdx {
						// MTO YTD: use lighter shades
						if mtoYTD <= 1 {
							fillColor = "#FFCCCC" // light red
						} else if mtoYTD <= 3 {
							fillColor = "#FFFFCC" // light yellow
						} else {
							fillColor = "#CCFFCC" // light greenn
						}
					} else {
						// MTO PY column: use darker shades
						if mtoPY <= 1 {
							fillColor = "#FF6666" // darker red
						} else if mtoPY <= 3 {
							fillColor = "#FFCC33" // darker yellow
						} else {
							fillColor = "#66FF66" // darker green
						}
					}
				}
				// if status is rundown or discontinued, override fill with gray shades;
				// otherwise preserve whatever fillColor was set above (MTO or default white)
				switch e.Status {
				case "Rundown":
					fillColor = "#D3D3D3" // light gray
				case "Discontinued":
					fillColor = "#A9A9A9" // dark gray
				}

				style, _ := f.NewStyle(&excelize.Style{
					// center alignment for all cells except Description
					Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
					// set borders for data cells
					Border: []excelize.Border{
						{Type: "left", Color: "000000", Style: 1},
						{Type: "right", Color: "000000", Style: 1},
						{Type: "top", Color: "000000", Style: 1},
						{Type: "bottom", Color: "000000", Style: 1},
					},
					Fill: excelize.Fill{
						Type:    "pattern",
						Color:   []string{fillColor},
						Pattern: 1,
					},
				})
				f.SetCellStyle(sh, cell, cell, style)
			}
			rowIdx[sh] = r + 1
		}

		// set column widths for better readability
		// Build column widths based on the headers we actually wrote (handles both PO and non-PO variants)
		colWidths := make(map[string]float64)
		for i, h := range headersOut {
			col, _ := excelize.ColumnNumberToName(i + 1)
			switch h {
			case "Item Code":
				colWidths[col] = 15
			case "QTY on Hand":
				colWidths[col] = 12
			case "Total QTY on PO":
				colWidths[col] = 15
			case "PO Num 1", "PO Num 2":
				colWidths[col] = 12
			case "QTY on PO 1", "QTY on PO 2":
				colWidths[col] = 12
			case "QTY on SO/BO":
				colWidths[col] = 15
			case "QTY Available":
				colWidths[col] = 15
			case "MTO YTD", "MTO PY":
				colWidths[col] = 10
			case "QTY Sold/Issued YTD", "QTY Sold/Issued PY":
				colWidths[col] = 18
			case "Class":
				colWidths[col] = 20
			case "Status":
				colWidths[col] = 15
			case "Occasion":
				colWidths[col] = 20
			case "Description":
				colWidths[col] = 40
			case "UPC":
				colWidths[col] = 15
			case "Foil":
				colWidths[col] = 10
			default:
				colWidths[col] = 12
			}
		}
		for _, sh := range []string{"Everyday", "Winter", "Spring"} {
			for col, width := range colWidths {
				f.SetColWidth(sh, col, col, width)
			}
		}

		// add autofilter to header row
		for _, sh := range []string{"Everyday", "Winter", "Spring"} {
			lastCol, _ := excelize.ColumnNumberToName(len(headersOut))
			f.AutoFilter(sh, fmt.Sprintf("A1:%s1", lastCol), nil)
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
