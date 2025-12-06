package hotsheet

import (
	"fmt"
	"strconv"
	"strings"
	"time"

	helpers "github.com/Fepozopo/bsc-hotsheet-update/helpers"
	"github.com/xuri/excelize/v2"
)

type Update struct {
	Hotsheet            string
	Sheet               string
	InventoryReport     string
	POReport            string
	BNReport            string
	SkuCol              string
	OnHandCol           string
	OnPOCol1            string
	OnPOCol2            string
	OnPOCol3            string
	OnPOColTotal        string
	OnSOBOCol           string
	YtdSoldIssuedCol    string
	PONumCol1           string
	PONumCol2           string
	PONumCol3           string
	AverageMonthlyCol   string
	BNYtdSoldCol        string
	BNAverageMonthlyCol string
}

// Update updates the hotsheet with stock and sales data from the report.
// It matches SKUs from the hotsheet with those in the report, retrieves
// relevant stock information, and updates
// the corresponding cells in the hotsheet.
//
// Parameters:
//   - product: A string representing the product name for logging purposes.
//   - occasion: A string representing the occasion for logging purposes.
//
// Returns:
//   - error: An error if any operation (e.g., file opening, reading, or writing)
//     fails during the update process.
func (u *Update) UpdateInventory(product, occasion string) error {
	logger, logFile, err := helpers.CreateLogger("inventory", product, occasion, "INFO")
	if err != nil {
		return fmt.Errorf("failed to create log file: %w", err)
	}
	defer logFile.Close()

	// Track whether any parse errors occurred so we can notify the user once
	var parseErrorOccurred bool

	// Get the current date
	currentDate := time.Now()

	// Get the current month, day, and last day of the current month
	currentMonth := float64(currentDate.Month())
	currentDay := float64(currentDate.Day())
	lastDay := float64(time.Date(currentDate.Year(), currentDate.Month()+1, 0, 0, 0, 0, 0, time.UTC).Day())

	// Calculate the fractional part of the month
	monthFloat := (currentMonth - 1) + (currentDay / lastDay)

	// Open the report workbook
	wbReport, err := excelize.OpenFile(u.InventoryReport)
	if err != nil {
		return fmt.Errorf("failed to open report file %s: %w", u.InventoryReport, err)
	}
	defer wbReport.Close()

	// Open the hotsheet workbook
	wbHotsheet, err := excelize.OpenFile(u.Hotsheet)
	if err != nil {
		return fmt.Errorf("failed to open hotsheet file %s: %w", u.Hotsheet, err)
	}
	defer wbHotsheet.Close()

	// Open the BN report workbook
	wbReportBN, err := excelize.OpenFile(u.BNReport)
	if err != nil {
		return fmt.Errorf("failed to open BN report file %s: %w", u.BNReport, err)
	}
	defer wbReportBN.Close()

	// Get the sheets
	wsReport := "Sheet1"
	wsHotsheet := u.Sheet
	wsReportBN := "Sheet1"

	// Get the rows
	rowsHotsheet, err := wbHotsheet.GetRows(wsHotsheet)
	if err != nil {
		return fmt.Errorf("failed to get rows from hotsheet file %s: %w", u.Hotsheet, err)
	}
	rowsReport, err := wbReport.GetRows(wsReport)
	if err != nil {
		return fmt.Errorf("failed to get rows from report file %s: %w", u.InventoryReport, err)
	}
	rowsReportBN, err := wbReportBN.GetRows(wsReportBN)
	if err != nil {
		return fmt.Errorf("failed to get rows from BN report file %s: %w", u.BNReport, err)
	}

	skuCol := "B"       // 'B' column index in wsReport
	onHandCol := "B"    // 'B' column index in wsReport
	onPOCol := "D"      // 'D' column index in wsReport
	onSOCol := "F"      // 'F' column index in wsReport
	onBOCol := "H"      // 'H' column index in wsReport
	ytdSoldCol := "L"   // 'L' column index in wsReport
	ytdIssuedCol := "N" // 'N' column index in wsReport

	bnSkuCol := "J"     // 'J' column index in wsReportBN
	bnYtdSoldCol := "O" // 'O' column index in wsReportBN

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

	// helper: parse numeric-like cell strings to int (handles commas and trailing dash)
	parseNum := func(s string) (int, error) {
		if s == "" {
			return 0, nil
		}
		s = strings.ReplaceAll(s, ",", "")
		if strings.HasSuffix(s, "-") {
			s = "-" + strings.TrimSuffix(s, "-")
		}
		f, err := strconv.ParseFloat(s, 64)
		if err != nil {
			return 0, err
		}
		return int(f), nil
	}

	// Build a map of inventory data keyed by SKU using the same offsets the original code used.
	type invData struct {
		onHand    int
		onPO      int
		onSO      int
		onBO      int
		ytdSold   int
		ytdIssued int
	}
	reportMap := make(map[string]invData)
	skuIdx := colToIndex(skuCol)
	onHandIdx := colToIndex(onHandCol)
	onPOIdx := colToIndex(onPOCol)
	onSOIdx := colToIndex(onSOCol)
	onBOIdx := colToIndex(onBOCol)
	ytdSoldIdx := colToIndex(ytdSoldCol)
	ytdIssuedIdx := colToIndex(ytdIssuedCol)

	// rowsReport is a slice of rows
	for rowNum := 1; rowNum < len(rowsReport)+1; rowNum++ {
		// get SKU at rowNum (which corresponds to rowsReport[rowNum-1])
		if rowNum-1 >= len(rowsReport) {
			continue
		}
		row := rowsReport[rowNum-1]
		if skuIdx >= len(row) {
			continue
		}
		sku := strings.TrimSpace(row[skuIdx])
		if sku == "" {
			continue
		}
		// valueLocation := rowNum + 2
		valueLocation := rowNum + 2
		if valueLocation-1 >= len(rowsReport) {
			continue
		}
		valRow := rowsReport[valueLocation-1]

		// safe-get helpers
		getCell := func(idx int) string {
			if idx < 0 || idx >= len(valRow) {
				return ""
			}
			return valRow[idx]
		}

		onHandStr := strings.ReplaceAll(getCell(onHandIdx), ",", "")
		onPOStr := strings.ReplaceAll(getCell(onPOIdx), ",", "")
		onSOStr := strings.ReplaceAll(getCell(onSOIdx), ",", "")
		onBOStr := strings.ReplaceAll(getCell(onBOIdx), ",", "")
		ytdSoldStr := strings.ReplaceAll(getCell(ytdSoldIdx), ",", "")
		ytdIssuedStr := strings.ReplaceAll(getCell(ytdIssuedIdx), ",", "")

		onHandInt, err1 := parseNum(onHandStr)
		onPOInt, err2 := parseNum(onPOStr)
		onSOInt, err3 := parseNum(onSOStr)
		onBOInt, err4 := parseNum(onBOStr)
		ytdSoldInt, err5 := parseNum(ytdSoldStr)
		ytdIssuedInt, err6 := parseNum(ytdIssuedStr)
		// if any parse fails, skip this row but log the problem
		if err1 != nil || err2 != nil || err3 != nil || err4 != nil || err5 != nil || err6 != nil {
			logger.Printf("Skipping SKU %s due to parse error: %v %v %v %v %v %v\n", sku, err1, err2, err3, err4, err5, err6)
			// Mark that a parse error occurred so we can notify the user after processing
			parseErrorOccurred = true
			continue
		}
		reportMap[sku] = invData{
			onHand:    onHandInt,
			onPO:      onPOInt,
			onSO:      onSOInt,
			onBO:      onBOInt,
			ytdSold:   ytdSoldInt,
			ytdIssued: ytdIssuedInt,
		}
	}

	// Build BN map using an offset of +1
	bnMap := make(map[string]int)
	bnSkuIdx := colToIndex(bnSkuCol)
	bnYtdSoldIdx := colToIndex(bnYtdSoldCol)
	for rowNum := 1; rowNum < len(rowsReportBN)+1; rowNum++ {
		if rowNum-1 >= len(rowsReportBN) {
			continue
		}
		row := rowsReportBN[rowNum-1]
		if bnSkuIdx >= len(row) {
			continue
		}
		sku := strings.TrimSpace(row[bnSkuIdx])
		if sku == "" {
			continue
		}
		// valueLocationBN := rowNum + 1
		valueLocationBN := rowNum + 1
		if valueLocationBN-1 >= len(rowsReportBN) {
			continue
		}
		valRow := rowsReportBN[valueLocationBN-1]
		getCell := func(idx int) string {
			if idx < 0 || idx >= len(valRow) {
				return ""
			}
			return valRow[idx]
		}
		bnYtdSoldStr := strings.ReplaceAll(getCell(bnYtdSoldIdx), ",", "")
		if strings.HasSuffix(bnYtdSoldStr, "-") {
			bnYtdSoldStr = "-" + strings.TrimSuffix(bnYtdSoldStr, "-")
		}
		if bnYtdSoldStr == "" {
			bnMap[sku] = 0
			continue
		}
		bnValFloat, err := strconv.ParseFloat(bnYtdSoldStr, 64)
		if err != nil {
			logger.Printf("Skipping BN SKU %s due to parse error: %v\n", sku, err)
			// Mark that a parse error occurred so we can notify the user after processing
			parseErrorOccurred = true
			continue
		}
		bnMap[sku] = int(bnValFloat)
	}

	// Progress bar
	var bar helpers.Bar
	bar.NewOption(int64(0), int64(len(rowsHotsheet)-1))

	// Iterate hotsheet and lookup in maps
	for rowWsHotsheet := 2; rowWsHotsheet < len(rowsHotsheet)+1; rowWsHotsheet++ {
		bar.Play(int64(rowWsHotsheet - 1))

		skuWsHotsheet, err := wbHotsheet.GetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.SkuCol, rowWsHotsheet)) // SKU column in wsHotsheet
		if err != nil {
			return fmt.Errorf("failed to get SKU from hotsheet file %s: %w", u.Hotsheet, err)
		}
		if skuWsHotsheet == "" {
			continue // Skip rows with no SKU in wsHotsheet
		}
		skuKey := strings.TrimSpace(skuWsHotsheet)

		data, ok := reportMap[skuKey]
		if !ok {
			// no match, clear PO columns
			continue
		}

		// Calculate onSOBO, ytdSoldIssued, and averageMonthly
		onSOBOInt := data.onSO + data.onBO
		ytdSoldIssuedInt := data.ytdSold + data.ytdIssued
		averageMonthly := float64(ytdSoldIssuedInt) / monthFloat

		// Update the hotsheet
		if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnHandCol, rowWsHotsheet), data.onHand); err != nil {
			return fmt.Errorf("failed to set onHand value in hotsheet file: %w", err)
		}
		if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOColTotal, rowWsHotsheet), data.onPO); err != nil {
			return fmt.Errorf("failed to set onPO value in hotsheet file: %w", err)
		}
		if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnSOBOCol, rowWsHotsheet), onSOBOInt); err != nil {
			return fmt.Errorf("failed to set onSOBO value in hotsheet file: %w", err)
		}
		if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.YtdSoldIssuedCol, rowWsHotsheet), ytdSoldIssuedInt); err != nil {
			return fmt.Errorf("failed to set ytdSoldIssued value in hotsheet file: %w", err)
		}
		if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.AverageMonthlyCol, rowWsHotsheet), averageMonthly); err != nil {
			return fmt.Errorf("failed to set averageMonthly value in hotsheet file: %w", err)
		}

		// Remove the old PO number and old PO quantities
		wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.PONumCol1, rowWsHotsheet), "")
		wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.PONumCol2, rowWsHotsheet), "")
		wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.PONumCol3, rowWsHotsheet), "")
		wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOCol1, rowWsHotsheet), "")
		wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOCol2, rowWsHotsheet), "")
		wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOCol3, rowWsHotsheet), "")

		// Remove possibly old BN values
		wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.BNYtdSoldCol, rowWsHotsheet), "")
		wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.BNAverageMonthlyCol, rowWsHotsheet), "")

		logger.Printf("Match found for SKU: %s | onHand: %d | onPO: %d | onSO: %d | onBO: %d | ytdSold: %d | ytdIssued: %d\n", skuWsHotsheet, data.onHand, data.onPO, data.onSO, data.onBO, data.ytdSold, data.ytdIssued)
		// BN handling: lookup in bnMap
		if bnVal, ok := bnMap[skuKey]; ok {
			bnYtdSoldInt := bnVal
			bnYtdSoldIssuedInt := (data.ytdSold + data.ytdIssued) - bnYtdSoldInt
			bnAverageMonthly := float64(bnYtdSoldIssuedInt) / monthFloat
			if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.BNYtdSoldCol, rowWsHotsheet), bnYtdSoldIssuedInt); err != nil {
				return fmt.Errorf("failed to set BN YTD Sold value in hotsheet file: %w", err)
			}
			if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.BNAverageMonthlyCol, rowWsHotsheet), bnAverageMonthly); err != nil {
				return fmt.Errorf("failed to set BN average monthly value in hotsheet file: %w", err)
			}
			logger.Printf("BN Match found for SKU: %s | BN YTD Sold: %d\n", skuWsHotsheet, bnYtdSoldInt)
		}
	}

	if err := wbHotsheet.UpdateLinkedValue(); err != nil {
		return fmt.Errorf("failed to update linked value in hotsheet file %s: %w", u.Hotsheet, err)
	}
	if err := wbHotsheet.Save(); err != nil {
		return fmt.Errorf("failed to save hotsheet file: %w", err)
	}

	// If any parse errors happened while reading the inventory/BN reports, let the user know to check the log file.
	if parseErrorOccurred {
		fmt.Printf("Parse errors occurred during inventory update; see log file: %s\n", logFile.Name())
	}

	bar.Finish()
	return nil
}

// UpdatePONumber updates the Purchase Order (PO) numbers in the hotsheet
// by matching SKUs between the hotsheet and the PO report.
// It logs each operation and writes the updated PO numbers back to the hotsheet.
//
// Parameters:
//   - product: A string representing the product name for logging purposes.
//   - occasion: A string representing the occasion for logging purposes.
//
// Returns:
//   - error: An error if any operation (e.g., file opening, reading, writing, or conversion)
//     fails during the PO number update process.
func (u *Update) UpdatePONumber(product, occasion string) error {
	logger, logFile, err := helpers.CreateLogger("PO", product, occasion, "INFO")
	if err != nil {
		return fmt.Errorf("failed to create log file: %w", err)
	}
	defer logFile.Close()

	// Track parse errors during PO processing so we can notify the user once
	var parseErrorOccurred bool

	// Open the report workbook
	wbReport, err := excelize.OpenFile(u.POReport)
	if err != nil {
		return fmt.Errorf("failed to open report file %s: %w", u.POReport, err)
	}
	defer wbReport.Close()

	// Open the hotsheet workbook
	wbHotsheet, err := excelize.OpenFile(u.Hotsheet)
	if err != nil {
		return fmt.Errorf("failed to open hotsheet file %s: %w", u.Hotsheet, err)
	}
	defer wbHotsheet.Close()

	// Get the sheets
	wsReport := "Sheet1"
	wsHotsheet := u.Sheet

	// Get the rows
	rowsHotsheet, err := wbHotsheet.GetRows(wsHotsheet)
	if err != nil {
		return fmt.Errorf("failed to get rows from hotsheet file %s: %w", u.Hotsheet, err)
	}
	rowsReport, err := wbReport.GetRows(wsReport)
	if err != nil {
		return fmt.Errorf("failed to get rows from report file %s: %w", u.POReport, err)
	}

	dataCol := "A"          // 'A' column index in wsReport
	onPOCol := "I"          // 'I' column index in wsReport
	onPOBackorderCol := "K" // 'K' column index in wsReport
	POStatusCol := "G"      // 'G' column index in wsReport

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

	// Build PO report map keyed by SKU. Preserve the original offsets used by the code:
	// valueLocation1 := rowWsReport + 1, valueLocation2 := rowWsReport + 2, valueLocation3 := rowWsReport + 3
	type poEntry struct {
		poNum1 string
		onPO1  string
		poNum2 string
		onPO2  string
		poNum3 string
		onPO3  string
	}
	poMap := make(map[string]poEntry)
	dataIdx := colToIndex(dataCol)
	onPOIdx := colToIndex(onPOCol)
	onPOBackIdx := colToIndex(onPOBackorderCol)
	poStatusIdx := colToIndex(POStatusCol)

	for rowNum := 1; rowNum < len(rowsReport)+1; rowNum++ {
		if rowNum-1 >= len(rowsReport) {
			continue
		}
		row := rowsReport[rowNum-1]
		if dataIdx >= len(row) {
			continue
		}
		sku := strings.TrimSpace(row[dataIdx])
		if sku == "" {
			continue
		}

		// Read the subsequent rows preserving offsets
		valueLocation1 := rowNum + 1
		valueLocation2 := rowNum + 2
		valueLocation3 := rowNum + 3

		getRow := func(vloc int) []string {
			if vloc-1 >= 0 && vloc-1 < len(rowsReport) {
				return rowsReport[vloc-1]
			}
			return nil
		}

		row1 := getRow(valueLocation1)
		row2 := getRow(valueLocation2)
		row3 := getRow(valueLocation3)

		getCell := func(r []string, idx int) string {
			if r == nil || idx < 0 || idx >= len(r) {
				return ""
			}
			return r[idx]
		}

		var poNum1, poNum2, poNum3, onPO1, onPO2, onPO3 string

		if row1 != nil {
			poNum1 = getCell(row1, dataIdx)
			poStatus1 := getCell(row1, poStatusIdx)
			if poStatus1 == "Back Order" {
				onPO1 = strings.ReplaceAll(getCell(row1, onPOBackIdx), ",", "")
			} else {
				onPO1 = strings.ReplaceAll(getCell(row1, onPOIdx), ",", "")
			}
		}
		if row2 != nil {
			poNum2 = getCell(row2, dataIdx)
			poStatus2 := getCell(row2, poStatusIdx)
			if poStatus2 == "Back Order" {
				onPO2 = strings.ReplaceAll(getCell(row2, onPOBackIdx), ",", "")
			} else {
				onPO2 = strings.ReplaceAll(getCell(row2, onPOIdx), ",", "")
			}
		}
		if row3 != nil {
			poNum3 = getCell(row3, dataIdx)
			poStatus3 := getCell(row3, poStatusIdx)
			if poStatus3 == "Back Order" {
				onPO3 = strings.ReplaceAll(getCell(row3, onPOBackIdx), ",", "")
			} else {
				onPO3 = strings.ReplaceAll(getCell(row3, onPOIdx), ",", "")
			}
		}

		poMap[sku] = poEntry{
			poNum1: strings.TrimSpace(poNum1),
			onPO1:  strings.TrimSpace(onPO1),
			poNum2: strings.TrimSpace(poNum2),
			onPO2:  strings.TrimSpace(onPO2),
			poNum3: strings.TrimSpace(poNum3),
			onPO3:  strings.TrimSpace(onPO3),
		}
	}

	// Progress bar
	var bar helpers.Bar
	bar.NewOption(int64(2), int64(len(rowsHotsheet)-1))

	for rowWsHotsheet := 2; rowWsHotsheet < len(rowsHotsheet)+1; rowWsHotsheet++ {
		bar.Play(int64(rowWsHotsheet - 1))

		skuWsHotsheet, err := wbHotsheet.GetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.SkuCol, rowWsHotsheet)) // SKU column in wsHotsheet
		if err != nil {
			return fmt.Errorf("failed to get SKU from hotsheet file %s: %w", u.Hotsheet, err)
		}
		if skuWsHotsheet == "" {
			continue // Skip rows with no SKU in wsHotsheet
		}

		// No reason to look for a match if onPO is 0
		onPO, err := wbHotsheet.GetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOColTotal, rowWsHotsheet)) // onPO column in wsHotsheet
		if err != nil {
			return fmt.Errorf("failed to get onPO value from hotsheet file %s: %w", u.Hotsheet, err)
		}
		if onPO == "0" {
			continue // Skip rows with no onPO value in wsHotsheet
		}

		skuKey := strings.TrimSpace(skuWsHotsheet)
		entry, ok := poMap[skuKey]
		if !ok {
			continue
		}

		// PO1
		var onPO1Int int
		poNum1Int := 0
		if entry.poNum1 != "" {
			if n, err := strconv.Atoi(entry.poNum1); err == nil {
				poNum1Int = n
			}
		}
		if entry.onPO1 != "" {
			_, err = fmt.Sscan(entry.onPO1, &onPO1Int)
			if err != nil {
				logger.Printf("failed to parse onPO1 for SKU %s: %v\n", skuKey, err)
				// mark that a parse error occurred and notify user later
				parseErrorOccurred = true
			}
		}
		if poNum1Int != 0 {
			if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.PONumCol1, rowWsHotsheet), poNum1Int); err != nil {
				return fmt.Errorf("failed to set PO number in hotsheet file %s: %w", u.Hotsheet, err)
			}
			if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOCol1, rowWsHotsheet), onPO1Int); err != nil {
				return fmt.Errorf("failed to set onPO value in hotsheet file %s: %w", u.Hotsheet, err)
			}
			logger.Printf("%s has a quantity of %d on PO number %d\n", skuWsHotsheet, onPO1Int, poNum1Int)
		}

		// PO2 (only if starts with "00")
		if strings.HasPrefix(entry.poNum2, "00") {
			var onPO2Int int
			poNum2Int := 0
			if entry.poNum2 != "" {
				if n, err := strconv.Atoi(entry.poNum2); err == nil {
					poNum2Int = n
				}
			}
			if entry.onPO2 != "" {
				_, err = fmt.Sscan(entry.onPO2, &onPO2Int)
				if err != nil {
					logger.Printf("failed to parse onPO2 for SKU %s: %v\n", skuKey, err)
					// mark that a parse error occurred and notify user later
					parseErrorOccurred = true
				}
			}
			if poNum2Int != 0 {
				if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.PONumCol2, rowWsHotsheet), poNum2Int); err != nil {
					return fmt.Errorf("failed to set PO number in hotsheet file %s: %w", u.Hotsheet, err)
				}
				if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOCol2, rowWsHotsheet), onPO2Int); err != nil {
					return fmt.Errorf("failed to set onPO value in hotsheet file %s: %w", u.Hotsheet, err)
				}
				logger.Printf("%s has a quantity of %d on PO number %d\n", skuWsHotsheet, onPO2Int, poNum2Int)
			}
		}

		// PO3 (only if starts with "00")
		if strings.HasPrefix(entry.poNum3, "00") {
			var onPO3Int int
			poNum3Int := 0
			if entry.poNum3 != "" {
				if n, err := strconv.Atoi(entry.poNum3); err == nil {
					poNum3Int = n
				}
			}
			if entry.onPO3 != "" {
				_, err = fmt.Sscan(entry.onPO3, &onPO3Int)
				if err != nil {
					logger.Printf("failed to parse onPO3 for SKU %s: %v\n", skuKey, err)
					// mark that a parse error occurred and notify user later
					parseErrorOccurred = true
				}
			}
			if poNum3Int != 0 {
				if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.PONumCol3, rowWsHotsheet), poNum3Int); err != nil {
					return fmt.Errorf("failed to set PO number in hotsheet file %s: %w", u.Hotsheet, err)
				}
				if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOCol3, rowWsHotsheet), onPO3Int); err != nil {
					return fmt.Errorf("failed to set onPO value in hotsheet file %s: %w", u.Hotsheet, err)
				}
				logger.Printf("%s has a quantity of %d on PO number %d\n", skuWsHotsheet, onPO3Int, poNum3Int)
			}
		}

	}

	if err := wbHotsheet.UpdateLinkedValue(); err != nil {
		return fmt.Errorf("failed to update linked value in hotsheet file %s: %w", u.Hotsheet, err)
	}
	if err := wbHotsheet.Save(); err != nil {
		return fmt.Errorf("failed to save hotsheet file: %w", err)
	}

	// If any parse errors happened while reading the PO report, let the user know to check the log file.
	if parseErrorOccurred {
		fmt.Printf("Parse errors occurred during PO update; see log file: %s\n", logFile.Name())
	}

	bar.Finish()
	return nil
}
