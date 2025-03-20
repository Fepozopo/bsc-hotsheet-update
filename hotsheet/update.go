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
	Hotsheet          string
	Sheet             string
	InventoryReport   string
	POReport          string
	SkuCol            string
	OnHandCol         string
	OnPOCol1          string
	OnPOCol2          string
	OnPOCol3          string
	OnPOColTotal      string
	OnSOBOCol         string
	YtdSoldIssuedCol  string
	PONumCol1         string
	PONumCol2         string
	PONumCol3         string
	AverageMonthlyCol string
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
				var onHand, onPO, onSO, onBO, ytdSold, ytdIssued string
				if onHand, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onHandCol, valueLocation)); err != nil {
					return fmt.Errorf("failed to get On Hand value from report file %s: %w", u.InventoryReport, err)
				}
				if onPO, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onPOCol, valueLocation)); err != nil {
					return fmt.Errorf("failed to get On PO value from report file %s: %w", u.InventoryReport, err)
				}
				if onSO, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onSOCol, valueLocation)); err != nil {
					return fmt.Errorf("failed to get On SO value from report file %s: %w", u.InventoryReport, err)
				}
				if onBO, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onBOCol, valueLocation)); err != nil {
					return fmt.Errorf("failed to get On BO value from report file %s: %w", u.InventoryReport, err)
				}
				if ytdSold, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", ytdSoldCol, valueLocation)); err != nil {
					return fmt.Errorf("failed to get YTD Sold value from report file %s: %w", u.InventoryReport, err)
				}
				if ytdIssued, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", ytdIssuedCol, valueLocation)); err != nil {
					return fmt.Errorf("failed to get YTD Issued value from report file %s: %w", u.InventoryReport, err)
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
				if _, err = fmt.Sscan(onHand, &onHandInt); err != nil {
					return fmt.Errorf("failed to convert onHand value to int: %w", err)
				}
				if _, err = fmt.Sscan(onPO, &onPOInt); err != nil {
					return fmt.Errorf("failed to convert onPO value to int: %w", err)
				}
				if _, err = fmt.Sscan(onSO, &onSOInt); err != nil {
					return fmt.Errorf("failed to convert onSO value to int: %w", err)
				}
				if _, err = fmt.Sscan(onBO, &onBOInt); err != nil {
					return fmt.Errorf("failed to convert onBO value to int: %w", err)
				}
				if _, err = fmt.Sscan(ytdSold, &ytdSoldInt); err != nil {
					return fmt.Errorf("failed to convert ytdSold value to int: %w", err)
				}
				if _, err = fmt.Sscan(ytdIssued, &ytdIssuedInt); err != nil {
					return fmt.Errorf("failed to convert ytdIssued value to int: %w", err)
				}

				// Calculate onSOBO, ytdSoldIssued, and averageMonthly
				onSOBOInt := onSOInt + onBOInt
				ytdSoldIssuedInt := ytdSoldInt + ytdIssuedInt
				averageMonthly := float64(ytdSoldIssuedInt) / monthFloat

				// Update the hotsheet
				if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnHandCol, rowWsHotsheet), onHandInt); err != nil {
					return fmt.Errorf("failed to set onHand value in hotsheet file: %w", err)
				}
				if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOColTotal, rowWsHotsheet), onPOInt); err != nil {
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
	wsReportPointer := 1    // Start pointer for wsReport

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
		onPO, err := wbHotsheet.GetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOColTotal, rowWsHotsheet)) // onPO column in wsHotsheet
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
				valueLocation1 := rowWsReport + 1
				valueLocation2 := rowWsReport + 2
				valueLocation3 := rowWsReport + 3

				// Get the PO numbers and the first PO amount
				var poNum1, poNum2, poNum3, onPO1, poStatus1 string
				if poNum1, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", dataCol, valueLocation1)); err != nil {
					return fmt.Errorf("failed to get PO number from report file %s: %w", u.POReport, err)
				}
				if poNum2, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", dataCol, valueLocation2)); err != nil {
					return fmt.Errorf("failed to get PO number from report file %s: %w", u.POReport, err)
				}
				if poNum3, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", dataCol, valueLocation3)); err != nil {
					return fmt.Errorf("failed to get PO number from report file %s: %w", u.POReport, err)
				}
				if poStatus1, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", POStatusCol, valueLocation1)); err != nil {
					return fmt.Errorf("failed to get PO status from report file %s: %w", u.POReport, err)
				}
				if poStatus1 == "Back Order" { // If the PO status column is 'Back Order', then get the backorder value
					if onPO1, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onPOBackorderCol, valueLocation1)); err != nil {
						return fmt.Errorf("failed to get backorder value from report file %s: %w", u.POReport, err)
					}
					// Remove the commas from the backorder value
					onPO1 = strings.ReplaceAll(onPO1, ",", "")
				} else { // If the PO status column is not 'Back Order', then get the onPO value
					if onPO1, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onPOCol, valueLocation1)); err != nil {
						return fmt.Errorf("failed to get onPO value from report file %s: %w", u.POReport, err)
					}
					// Remove the commas from the onPO value
					onPO1 = strings.ReplaceAll(onPO1, ",", "")
				}

				// Convert the values to an integer. On PO values have to use fmt.Sscanf
				var onPO1Int int
				poNum1Int, err := strconv.Atoi(poNum1)
				if err != nil {
					return fmt.Errorf("failed to convert PO number to integer: %w", err)
				}
				_, err = fmt.Sscan(onPO1, &onPO1Int)
				if err != nil {
					return fmt.Errorf("failed to convert onPO value to integer: %w", err)
				}

				// Update the PO number in the hotsheet
				if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.PONumCol1, rowWsHotsheet), poNum1Int); err != nil {
					return fmt.Errorf("failed to set PO number in hotsheet file %s: %w", u.Hotsheet, err)
				}
				// Update the onPO value in the hotsheet
				if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOCol1, rowWsHotsheet), onPO1Int); err != nil {
					return fmt.Errorf("failed to set onPO value in hotsheet file %s: %w", u.Hotsheet, err)
				}
				logger.Printf("%s has a quantity of %d on PO number %d\n", skuWsHotsheet, onPO1Int, poNum1Int)

				// If poNum2 starts with "00"
				if strings.HasPrefix(poNum2, "00") {
					// Get the PO numbers and the second PO amount
					var onPO2, poStatus2 string
					if poStatus2, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", POStatusCol, valueLocation2)); err != nil {
						return fmt.Errorf("failed to get PO status from report file %s: %w", u.POReport, err)
					}
					if poStatus2 == "Back Order" { // If the PO status column is 'Back Order', then get the backorder value
						if onPO2, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onPOBackorderCol, valueLocation2)); err != nil {
							return fmt.Errorf("failed to get backorder value from report file %s: %w", u.POReport, err)
						}
						// Remove the comma from the backorder value
						onPO2 = strings.ReplaceAll(onPO2, ",", "")
					} else { // If the PO status column is not 'Back Order', then get the onPO value
						if onPO2, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onPOCol, valueLocation2)); err != nil {
							return fmt.Errorf("failed to get onPO value from report file %s: %w", u.POReport, err)
						}
						// Remove the comma from the backorder value
						onPO2 = strings.ReplaceAll(onPO2, ",", "")
					}
					// Convert the values to an integer
					var onPO2Int int
					poNum2Int, err := strconv.Atoi(poNum2)
					if err != nil {
						return fmt.Errorf("failed to convert PO number to integer: %w", err)
					}
					_, err = fmt.Sscan(onPO2, &onPO2Int)
					if err != nil {
						return fmt.Errorf("failed to convert onPO value to integer: %w", err)
					}
					// Update the PO number in the hotsheet
					if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.PONumCol2, rowWsHotsheet), poNum2Int); err != nil {
						return fmt.Errorf("failed to set PO number in hotsheet file %s: %w", u.Hotsheet, err)
					}
					// Update the onPO value in the hotsheet
					if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOCol2, rowWsHotsheet), onPO2Int); err != nil {
						return fmt.Errorf("failed to set onPO value in hotsheet file %s: %w", u.Hotsheet, err)
					}
					logger.Printf("%s has a quantity of %d on PO number %d\n", skuWsHotsheet, onPO2Int, poNum2Int)
				}

				// If poNum3 starts with "00"
				if strings.HasPrefix(poNum3, "00") {
					// Get the PO numbers and the second PO amount
					var onPO3, poStatus3 string
					if poStatus3, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", POStatusCol, valueLocation3)); err != nil {
						return fmt.Errorf("failed to get PO status from report file %s: %w", u.POReport, err)
					}
					if poStatus3 == "Back Order" { // If the PO status column is 'Back Order', then get the backorder value
						if onPO3, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onPOBackorderCol, valueLocation3)); err != nil {
							return fmt.Errorf("failed to get backorder value from report file %s: %w", u.POReport, err)
						}
						// Remove the comma from the backorder value
						onPO3 = strings.ReplaceAll(onPO3, ",", "")
					} else { // If the PO status column is not 'Back Order', then get the onPO value
						if onPO3, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onPOCol, valueLocation3)); err != nil {
							return fmt.Errorf("failed to get onPO value from report file %s: %w", u.POReport, err)
						}
						// Remove the comma from the onPO value
						onPO3 = strings.ReplaceAll(onPO3, ",", "")
					}
					// Convert the values to an integer
					poNum3Int, err := strconv.Atoi(poNum3)
					if err != nil {
						return fmt.Errorf("failed to convert PO number to integer: %w", err)
					}
					var onPO3Int int
					_, err = fmt.Sscan(onPO3, &onPO3Int)
					if err != nil {
						return fmt.Errorf("failed to convert onPO value to integer: %w", err)
					}
					// Update the PO number in the hotsheet
					if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.PONumCol3, rowWsHotsheet), poNum3Int); err != nil {
						return fmt.Errorf("failed to set PO number in hotsheet file %s: %w", u.Hotsheet, err)
					}
					// Update the onPO value in the hotsheet
					if err := wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", u.OnPOCol3, rowWsHotsheet), onPO3Int); err != nil {
						return fmt.Errorf("failed to set onPO value in hotsheet file %s: %w", u.Hotsheet, err)
					}
					logger.Printf("%s has a quantity of %d on PO number %d\n", skuWsHotsheet, onPO3Int, poNum3Int)
				}

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
