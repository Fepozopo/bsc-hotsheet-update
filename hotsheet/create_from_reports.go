package hotsheet

import (
	"fmt"
	"path/filepath"
	"sort"
	"strconv"
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

	// entry represents a single inventory item
	type entry struct {
		SKU         string
		ProductLine string
		ClassDesc   string
		Status      string
		OnHand      int
		OnPO        int
		// Per-PO details from the PO report
		PONum1         string
		OnPO1          int
		PONum2         string
		OnPO2          int
		PONum3         string
		OnPO3          int
		OnSO           int
		OnBO           int
		TotalAvailable int
		YTDSold        int
		YTDIssued      int
		SoldPY         int
		IssuedPY       int
		Foil           string
		Occasion       string
		Description    string
		UPC            string
	}

	invMap := make(map[string]*entry)

	// helper: convert column letter(s) to zero-based index (A -> 0, B -> 1, ...)
	colToIndex := func(col string) int {
		col = strings.ToUpper(col)
		idx := 0
		for i := 0; i < len(col); i++ {
			idx *= 26
			idx += int(col[i]-'A') + 1
		}
		return idx - 1
	}

	// parse numbers permissively
	parseInt := func(s string) int {
		s = strings.TrimSpace(strings.ReplaceAll(s, ",", ""))
		if s == "" {
			return 0
		}
		// handle trailing dash like "-" meaning negative
		if strings.HasSuffix(s, "-") {
			s = "-" + strings.TrimSuffix(s, "-")
		}
		v, err := strconv.Atoi(s)
		if err != nil {
			f, err2 := strconv.ParseFloat(s, 64)
			if err2 != nil {
				return 0
			}
			return int(f)
		}
		return v
	}

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

	// helper to read a cell by 1-based row and column index
	getCellAt := func(rows [][]string, rowNum int, colIdx int) string {
		if rowNum-1 < 0 || rowNum-1 >= len(rows) {
			return ""
		}
		row := rows[rowNum-1]
		if colIdx < 0 || colIdx >= len(row) {
			return ""
		}
		return strings.TrimSpace(row[colIdx])
	}

	// detect run-date row (examples like "4/6/2026  5:06:06 PM")
	isRunDate := func(s string) bool {
		s = strings.TrimSpace(s)
		if s == "" {
			return false
		}
		// Consider it a run-date row if it contains either a slash or a colon
		if strings.Contains(s, "/") || strings.Contains(s, ":") {
			return true
		}
		return false
	}

	startRow := 2
	for r := startRow; ; r += 3 {
		// SKU row is at column B, row r
		if r-1 >= len(invRows) {
			break
		}
		sku := getCellAt(invRows, r, skuIdx)
		if sku == "" {
			break
		}
		// if the SKU cell looks like a run date, stop parsing
		if isRunDate(sku) {
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
				poNumCol1 := "F"
				poNumCol2 := "H"
				poNumCol3 := "J"

				colToIndexPO := func(col string) int {
					col = strings.ToUpper(col)
					idx := 0
					for i := 0; i < len(col); i++ {
						idx *= 26
						idx += int(col[i]-'A') + 1
					}
					return idx - 1
				}

				dataIdx := colToIndexPO(dataCol)
				onPOIdx := colToIndexPO(onPOCol)
				onPOBackIdx := colToIndexPO(onPOBackorderCol)
				poStatusIdx := colToIndexPO(poStatusCol)
				poNum1Idx := colToIndexPO(poNumCol1)
				poNum2Idx := colToIndexPO(poNumCol2)
				poNum3Idx := colToIndexPO(poNumCol3)

				getRow := func(vloc int) []string {
					if vloc-1 >= 0 && vloc-1 < len(poRows) {
						return poRows[vloc-1]
					}
					return nil
				}

				getCell := func(r []string, idx int) string {
					if r == nil || idx < 0 || idx >= len(r) {
						return ""
					}
					return r[idx]
				}

				for rowNum := 1; rowNum < len(poRows)+1; rowNum++ {
					row := poRows[rowNum-1]
					if dataIdx >= len(row) {
						continue
					}
					sku := strings.TrimSpace(row[dataIdx])
					if sku == "" {
						continue
					}

					row1 := getRow(rowNum + 1)
					row2 := getRow(rowNum + 2)
					row3 := getRow(rowNum + 3)

					// ensure entry exists
					e, ok := invMap[sku]
					if !ok {
						e = &entry{SKU: sku}
						invMap[sku] = e
					}

					assignPO := func(poNum string, qty int) {
						if poNum == "" && qty == 0 {
							return
						}
						if e.PONum1 == "" {
							e.PONum1 = poNum
							e.OnPO1 = qty
							return
						}
						if e.PONum2 == "" {
							e.PONum2 = poNum
							e.OnPO2 = qty
							return
						}
						if e.PONum3 == "" {
							e.PONum3 = poNum
							e.OnPO3 = qty
							return
						}
						// fallback: accumulate into OnPO1
						e.OnPO1 += qty
					}

					if row1 != nil {
						status := strings.TrimSpace(getCell(row1, poStatusIdx))
						qty := 0
						if strings.EqualFold(status, "Back Order") {
							qty = parseInt(getCell(row1, onPOBackIdx))
						} else {
							qty = parseInt(getCell(row1, onPOIdx))
						}
						poNum := strings.TrimSpace(getCell(row1, poNum1Idx))
						assignPO(poNum, qty)
					}
					if row2 != nil {
						status := strings.TrimSpace(getCell(row2, poStatusIdx))
						qty := 0
						if strings.EqualFold(status, "Back Order") {
							qty = parseInt(getCell(row2, onPOBackIdx))
						} else {
							qty = parseInt(getCell(row2, onPOIdx))
						}
						poNum := strings.TrimSpace(getCell(row2, poNum2Idx))
						assignPO(poNum, qty)
					}
					if row3 != nil {
						status := strings.TrimSpace(getCell(row3, poStatusIdx))
						qty := 0
						if strings.EqualFold(status, "Back Order") {
							qty = parseInt(getCell(row3, onPOBackIdx))
						} else {
							qty = parseInt(getCell(row3, onPOIdx))
						}
						poNum := strings.TrimSpace(getCell(row3, poNum3Idx))
						assignPO(poNum, qty)
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
			pl = "UNKNOWN"
		}
		productGroups[pl] = append(productGroups[pl], e)
	}

	// prepare occasion token lists (uppercase)
	everTokens := []string{"SYMPATHY", "PET SYMPATHY", "LOVE", "ENCOURAGEMENT", "THANK YOU", "BIRTHDAY", "BLANK", "BAPTISM-COMMUNION", "BABY", "CONGRATULATIONS", "NEW HOME", "CAMP", "CANCER", "THINKING OF YOU", "GET WELL", "KID BIRTHDAY", "ALL OCCASION", "FRIENDSHIP", "MENOPAUSE", "MISS YOU", "SORRY", "TEACHER APPRECIATION", "WEDDING ANNIVERSARY"}
	winterTokens := []string{"CHRISTMAS", "HALLOWEEN", "THANKSGIVING", "VETERAN'S DAY", "VETERANS DAY"}
	springTokens := []string{"EASTER", "FATHER'S DAY", "FATHERS DAY", "GRADUATION", "INDEPENDENCE DAY", "MOTHER'S DAY", "MOTHERS DAY", "ST. PATRICK'S DAY", "ST PATRICKS DAY", "VALENTINE'S DAY", "VALENTINES DAY"}

	mapOccasion := func(occ string) string {
		o := strings.ToUpper(strings.TrimSpace(occ))
		if o == "" {
			return "Everyday"
		}
		for _, t := range springTokens {
			if strings.Contains(o, t) {
				return "Spring"
			}
		}
		for _, t := range winterTokens {
			if strings.Contains(o, t) {
				return "Winter"
			}
		}
		// default/explicit everyday matches
		for _, t := range everTokens {
			if strings.Contains(o, t) {
				return "Everyday"
			}
		}
		return "Everyday"
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
		logger.Printf("Generated hotsheet: %s", outPath)
	}

	return outputs, nil
}

// simple sanitizer for file names
func sanitizeFileName(s string) string {
	s = strings.TrimSpace(s)
	if s == "" {
		return "unknown"
	}
	// replace path separators and colons
	s = strings.ReplaceAll(s, string(filepath.Separator), "_")
	s = strings.ReplaceAll(s, ":", "")
	return s
}
