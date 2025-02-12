package helpers

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

type UpdateSales struct {
	hotsheet string
	sheet    string
	report   string
	skuCol   string
	ytdCol   string
}

// UpdateSales updates the specified hotsheet Excel file with sales data from the report file.
//
// Parameters:
//   - product: A string representing the product name to be included in the log file name.
//   - occasion: A string representing the occasion type (e.g. "everyday", "holiday") to be included in the log file name.
//
// Returns:
//   - An error if the operation fails.
func (us *UpdateSales) UpdateSales(product, occasion string) error {
	logger, logFile, err := CreateLogger("UpdateSales", product, occasion, "INFO")
	if err != nil {
		return fmt.Errorf("failed to create log file: %w", err)
	}
	defer logFile.Close()

	// Open the report workbook
	wbReport, err := excelize.OpenFile(us.report)
	if err != nil {
		return fmt.Errorf("failed to open report file %s: %w", us.report, err)
	}
	defer wbReport.Close()

	// Open the hotsheet workbook
	wbHotsheet, err := excelize.OpenFile(us.hotsheet)
	if err != nil {
		return fmt.Errorf("failed to open hotsheet file %s: %w", us.hotsheet, err)
	}
	defer wbHotsheet.Close()

	// Get the sheets
	wsReport := "Sheet1"
	wsHotsheet := us.sheet

	// Get the rows
	rowsHotsheet, err := wbHotsheet.GetRows(wsHotsheet)
	if err != nil {
		return fmt.Errorf("failed to get rows from hotsheet file %s: %w", us.hotsheet, err)
	}
	rowsReport, err := wbReport.GetRows(wsReport)
	if err != nil {
		return fmt.Errorf("failed to get rows from report file %s: %w", us.report, err)
	}

	skuCol := "A"        // 'A' column index in wsReport
	ytdCol := "S"        // 'S' column index in wsReport
	kitCol := "J"        // 'J' column index in wsReport
	wsReportPointer := 1 // Start pointer for wsReport

	// Progress bar
	var bar Bar
	bar.NewOption(int64(2), int64(len(rowsHotsheet)))

	for rowWsHotsheet := 2; rowWsHotsheet < len(rowsHotsheet)+1; rowWsHotsheet++ {
		skuWsHotsheet, err := wbHotsheet.GetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.skuCol, rowWsHotsheet))
		if err != nil {
			return fmt.Errorf("failed to get SKU from hotsheet file %s: %w", us.hotsheet, err)
		}

		if skuWsHotsheet == "" {
			continue // Skip rows with no SKU in wsHotsheet
		}

		for rowWsReport := wsReportPointer; rowWsReport < len(rowsReport); rowWsReport++ {
			skuWsReport, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", skuCol, rowWsReport)) // SKU in column 'A' in wsReport
			if err != nil {
				return fmt.Errorf("failed to get SKU from report file %s: %w", us.report, err)
			}

			if skuWsReport == "" {
				continue
			}

			logger.Printf("Comparing Hotsheet SKU: [%d] - '%s' with Report SKU: [%d] - '%s'\n", rowWsHotsheet, skuWsHotsheet, rowWsReport, skuWsReport)
			if strings.TrimSpace(skuWsHotsheet) == strings.TrimSpace(skuWsReport) {
				isKit, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", kitCol, rowWsReport+1))
				if err != nil {
					return fmt.Errorf("failed to get isKit from report file %s: %w", us.report, err)
				}
				ytdValue, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", ytdCol, rowWsReport+2))
				if err != nil {
					return fmt.Errorf("failed to get ytdValue from report file %s: %w", us.report, err)
				}

				// Remove commas from ytdValue
				ytdValue = strings.ReplaceAll(ytdValue, ",", "") // Remove commas

				// Convert ytdValue to int
				var ytdValueInt int
				_, err = fmt.Sscan(ytdValue, &ytdValueInt)
				if err != nil {
					return fmt.Errorf("failed to convert ytdValue to int: %w", err)
				}

				if isKit == "Kit" {
					if strings.Contains(skuWsReport, "20-") || strings.Contains(skuWsReport, "21-") ||
						strings.Contains(skuWsReport, "22-") || strings.Contains(skuWsReport, "24-") ||
						strings.Contains(skuWsReport, "20F-") || strings.Contains(skuWsReport, "22F-") ||
						strings.Contains(skuWsReport, "24F-") {
						// Update (ytd) in wsHotsheet
						wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.ytdCol, rowWsHotsheet), ytdValueInt)
					} else {
						ytdValue10x := ytdValueInt * 10
						wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.ytdCol, rowWsHotsheet), ytdValue10x)
					}
				} else {
					// Update (ytd) in wsHotsheet
					wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.ytdCol, rowWsHotsheet), ytdValueInt)
				}

				logger.Printf("Match found for SKU: %s | YTD: %d\n", skuWsHotsheet, ytdValueInt)
				wsReportPointer = rowWsReport + 1
				bar.Play(int64(rowWsHotsheet))
				break // Move to the next row in wsHotsheet once a match is found
			}
		}
	}

	if err := wbHotsheet.UpdateLinkedValue(); err != nil {
		return fmt.Errorf("failed to update linked value in hotsheet file %s: %w", us.hotsheet, err)
	}
	if err := wbHotsheet.Save(); err != nil {
		return fmt.Errorf("failed to save hotsheet file %s: %w", us.hotsheet, err)
	}

	bar.Finish()
	return nil
}
