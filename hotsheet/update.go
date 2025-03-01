package hotsheet

import (
	"fmt"
	"strconv"
	"strings"

	helpers "github.com/Fepozopo/bsc-hotsheet-update/helpers"
	"github.com/xuri/excelize/v2"
)

type Update struct {
	Hotsheet         string
	Sheet            string
	InventoryReport  string
	POReport         string
	SkuCol           string
	OnHandCol        string
	OnPOCol          string
	OnSOBOCol        string
	YtdSoldIssuedCol string
	PONumCol         string
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
func (u *Update) Update(product, occasion string) error {
	logger, logFile, err := helpers.CreateLogger("hotsheet", product, occasion, "INFO")
	if err != nil {
		return fmt.Errorf("failed to create log file: %w", err)
	}
	defer logFile.Close()

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
		return fmt.Errorf("failed to get rows from report file %s: %w", u.InventoryReport, err)
	}

	skuCol := "B"        // 'B' column index in wsReport
	onHandCol := "B"     // 'B' column index in wsReport
	onPOCol := "D"       // 'D' column index in wsReport
	onSOCol := "F"       // 'F' column index in wsReport
	onBOCol := "H"       // 'H' column index in wsReport
	ytdSoldCol := "L"    // 'L' column index in wsReport
	ytdIssuedCol := "N"  // 'N' column index in wsReport
	wsReportPointer := 1 // Start pointer for wsReport

	// Progress bar
	var bar helpers.Bar
	bar.NewOption(int64(2), int64(len(rowsHotsheet)))

	for rowWsHotsheet := 2; rowWsHotsheet < len(rowsHotsheet)+1; rowWsHotsheet++ {

		skuWsHotsheet, err := wbHotsheet.GetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.SkuCol, rowWsHotsheet)) // SKU column in wsHotsheet
		if err != nil {
			return fmt.Errorf("failed to get SKU from hotsheet file %s: %w", u.Hotsheet, err)
		}
		if skuWsHotsheet == "" {
			continue // Skip rows with no SKU in wsHotsheet
		}

		for rowWsReport := wsReportPointer; rowWsReport < len(rowsReport)+1; rowWsReport++ {
			skuWsReport, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", skuCol, rowWsReport)) // SKU in column 'B' in wsReport
			if err != nil {
				return fmt.Errorf("failed to get SKU from report file %s: %w", u.InventoryReport, err)
			}
			if skuWsReport == "" {
				continue
			}

			logger.Printf("Comparing Hotsheet SKU: [%d]-'%s' with Report SKU: [%d]-'%s'\n", rowWsHotsheet, skuWsHotsheet, rowWsReport, skuWsReport)
			if strings.TrimSpace(skuWsHotsheet) == strings.TrimSpace(skuWsReport) {
				valueLocation := rowWsReport + 2

				// Get the values for the current SKU in wsReport
				onHand, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onHandCol, valueLocation))
				if err != nil {
					return fmt.Errorf("failed to get on_hand value from report file %s: %w", u.InventoryReport, err)
				}
				onPO, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onPOCol, valueLocation))
				if err != nil {
					return fmt.Errorf("failed to get on_po value from report file %s: %w", u.InventoryReport, err)
				}
				onSO, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onSOCol, valueLocation))
				if err != nil {
					return fmt.Errorf("failed to get on_so value from report file %s: %w", u.InventoryReport, err)
				}
				onBO, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onBOCol, valueLocation))
				if err != nil {
					return fmt.Errorf("failed to get on_bo value from report file %s: %w", u.InventoryReport, err)
				}
				ytdSold, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", ytdSoldCol, valueLocation))
				if err != nil {
					return fmt.Errorf("failed to get ytd_sold value from report file %s: %w", u.InventoryReport, err)
				}
				ytdIssued, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", ytdIssuedCol, valueLocation))
				if err != nil {
					return fmt.Errorf("failed to get ytd_issued value from report file %s: %w", u.InventoryReport, err)
				}

				// Replace commas with empty strings
				onHand = strings.ReplaceAll(onHand, ",", "")
				onPO = strings.ReplaceAll(onPO, ",", "")
				onSO = strings.ReplaceAll(onSO, ",", "")
				onBO = strings.ReplaceAll(onBO, ",", "")
				ytdSold = strings.ReplaceAll(ytdSold, ",", "")
				ytdIssued = strings.ReplaceAll(ytdIssued, ",", "")

				// Convert the values to integers
				var onHandInt, onPOInt, onSOInt, onBOInt, ytdSoldInt, ytdIssuedInt int
				_, err = fmt.Sscan(onHand, &onHandInt)
				if err != nil {
					return fmt.Errorf("failed to convert onHand value to int: %w", err)
				}
				_, err = fmt.Sscan(onPO, &onPOInt)
				if err != nil {
					return fmt.Errorf("failed to convert onPO value to int: %w", err)
				}
				_, err = fmt.Sscan(onSO, &onSOInt)
				if err != nil {
					return fmt.Errorf("failed to convert onSO value to int: %w", err)
				}
				_, err = fmt.Sscan(onBO, &onBOInt)
				if err != nil {
					return fmt.Errorf("failed to convert onBO value to int: %w", err)
				}
				_, err = fmt.Sscan(ytdSold, &ytdSoldInt)
				if err != nil {
					return fmt.Errorf("failed to convert ytdSold value to int: %w", err)
				}
				_, err = fmt.Sscan(ytdIssued, &ytdIssuedInt)
				if err != nil {
					return fmt.Errorf("failed to convert ytdIssued value to int: %w", err)
				}

				// Calculate onSOBO and ytdSoldIssued
				onSOBOInt := onSOInt + onBOInt
				ytdSoldIssuedInt := ytdSoldInt + ytdIssuedInt

				// Update the hotsheet
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnHandCol, rowWsHotsheet), onHandInt)
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOCol, rowWsHotsheet), onPOInt)
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnSOBOCol, rowWsHotsheet), onSOBOInt)
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.YtdSoldIssuedCol, rowWsHotsheet), ytdSoldIssuedInt)

				// Remove the old PO number
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.PONumCol, rowWsHotsheet), "")

				logger.Printf("Match found for SKU: %s | onHand: %d | onPO: %d | onSO: %d | onBO: %d | ytdSold: %d | ytdIssued: %d\n", skuWsHotsheet, onHandInt, onPOInt, onSOInt, onBOInt, ytdSoldInt, ytdIssuedInt)
				wsReportPointer = rowWsReport + 1
				bar.Play(int64(rowWsHotsheet))
				break // Move to the next row in wsHotsheet once a match is found

			}
		}
	}

	if err := wbHotsheet.UpdateLinkedValue(); err != nil {
		return fmt.Errorf("failed to update linked value in hotsheet file %s: %w", u.Hotsheet, err)
	}
	if err := wbHotsheet.Save(); err != nil {
		return fmt.Errorf("failed to save hotsheet file: %w", err)
	}

	bar.Finish()
	return nil
}

func (u *Update) UpdatePONumber(product, occasion string) error {
	logger, logFile, err := helpers.CreateLogger("PO", product, occasion, "INFO")
	if err != nil {
		return fmt.Errorf("failed to create log file: %w", err)
	}
	defer logFile.Close()

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

	dataCol := "A"       // 'A' column index in wsReport
	wsReportPointer := 1 // Start pointer for wsReport

	// Progress bar
	var bar helpers.Bar
	bar.NewOption(int64(2), int64(len(rowsHotsheet)))

	for rowWsHotsheet := 2; rowWsHotsheet < len(rowsHotsheet)+1; rowWsHotsheet++ {

		skuWsHotsheet, err := wbHotsheet.GetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.SkuCol, rowWsHotsheet)) // SKU column in wsHotsheet
		if err != nil {
			return fmt.Errorf("failed to get SKU from hotsheet file %s: %w", u.Hotsheet, err)
		}
		if skuWsHotsheet == "" {
			continue // Skip rows with no SKU in wsHotsheet
		}

		// No reason to look for a match if onPO is 0
		onPO, err := wbHotsheet.GetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOCol, rowWsHotsheet)) // onPO column in wsHotsheet
		if err != nil {
			return fmt.Errorf("failed to get onPO value from hotsheet file %s: %w", u.Hotsheet, err)
		}
		if onPO == "0" {
			continue // Skip rows with no onPO value in wsHotsheet
		}

		for rowWsReport := wsReportPointer; rowWsReport < len(rowsReport)+1; rowWsReport++ {
			skuWsReport, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", dataCol, rowWsReport)) // SKU in column 'A' in wsReport
			if err != nil {
				return fmt.Errorf("failed to get SKU from report file %s: %w", u.POReport, err)
			}
			if skuWsReport == "" {
				continue
			}

			logger.Printf("Comparing Hotsheet SKU: [%d]-'%s' with Report SKU: [%d]-'%s'\n", rowWsHotsheet, skuWsHotsheet, rowWsReport, skuWsReport)
			if strings.TrimSpace(skuWsHotsheet) == strings.TrimSpace(skuWsReport) {
				valueLocation := rowWsReport + 1

				// Get the PO number
				poNum, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", dataCol, valueLocation))
				if err != nil {
					return fmt.Errorf("failed to get PO number from report file %s: %w", u.POReport, err)
				}

				// Convert the value to an integer
				poNumInt, err := strconv.Atoi(poNum)
				if err != nil {
					return fmt.Errorf("failed to convert PO number to integer: %w", err)
				}

				// Update the PO number in the hotsheet
				if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.PONumCol, rowWsHotsheet), poNumInt); err != nil {
					return fmt.Errorf("failed to set PO number in hotsheet file %s: %w", u.Hotsheet, err)
				}

				logger.Printf("%s is on PO number %d\n", skuWsHotsheet, poNumInt)
				wsReportPointer = rowWsReport + 1
				bar.Play(int64(rowWsHotsheet))
				break // Move to the next row in wsHotsheet once a match is found
			}
		}
	}

	if err := wbHotsheet.UpdateLinkedValue(); err != nil {
		return fmt.Errorf("failed to update linked value in hotsheet file %s: %w", u.Hotsheet, err)
	}
	if err := wbHotsheet.Save(); err != nil {
		return fmt.Errorf("failed to save hotsheet file: %w", err)
	}

	bar.Finish()
	return nil
}
