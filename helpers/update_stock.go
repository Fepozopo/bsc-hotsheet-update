package helpers

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

type UpdateStock struct {
	hotsheet  string
	sheet     string
	report    string
	skuCol    string
	onHandCol string
	onPOCol   string
	onSOBOCol string
}

func (us *UpdateStock) UpdateStock(product, occasion string) error {
	logger, logFile, err := CreateLogger("UpdateStock", product, occasion, "INFO")
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
	onHandCol := "K"     // 'K' column index in wsReport
	onPOCol := "M"       // 'M' column index in wsReport
	onSOCol := "P"       // 'P' column index in wsReport
	onBOCol := "T"       // 'T' column index in wsReport
	wsReportPointer := 1 // Start pointer for wsReport

	// Progress bar
	var bar Bar
	bar.NewOption(int64(2), int64(len(rowsHotsheet)))

	for rowWsHotsheet := 2; rowWsHotsheet < len(rowsHotsheet)+1; rowWsHotsheet++ {

		skuWsHotsheet, err := wbHotsheet.GetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.skuCol, rowWsHotsheet)) // SKU column in wsHotsheet
		if err != nil {
			return fmt.Errorf("failed to get SKU from hotsheet file %s: %w", us.hotsheet, err)
		}

		if skuWsHotsheet == "" {
			continue // Skip rows with no SKU in wsHotsheet
		}

		for rowWsReport := wsReportPointer; rowWsReport < len(rowsReport)+1; rowWsReport++ {
			skuWsReport, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", skuCol, rowWsReport)) // SKU in column 'A' in wsReport
			if err != nil {
				return fmt.Errorf("failed to get SKU from report file %s: %w", us.report, err)
			}

			if skuWsReport == "" {
				continue
			}

			// Skip the first sku that contains "MPN48"
			if skuWsReport == "MPN48 BOX  HUMMINGBIRD BOX OF 10 NOTECARD" {
				continue
			}

			logger.Printf("Comparing Hotsheet SKU: [%d] - '%s' with Report SKU: [%d] - '%s'\n", rowWsHotsheet, skuWsHotsheet, rowWsReport, skuWsReport)
			if strings.Contains(skuWsReport, skuWsHotsheet) {
				testValueLocation, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onHandCol, rowWsReport+2))
				if err != nil {
					return fmt.Errorf("failed to get test value from report file %s: %w", us.report, err)
				}

				var valueLocation int
				if testValueLocation == "" {
					valueLocation = 1
				} else {
					valueLocation = 2
				}

				// Get the values (on_hand), (on_po), and (on_so, on_bo) in wsReport
				onHand, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onHandCol, rowWsReport+valueLocation))
				if err != nil {
					return fmt.Errorf("failed to get on_hand value from report file %s: %w", us.report, err)
				}
				onPO, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onPOCol, rowWsReport+valueLocation))
				if err != nil {
					return fmt.Errorf("failed to get on_po value from report file %s: %w", us.report, err)
				}
				onSO, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onSOCol, rowWsReport+valueLocation))
				if err != nil {
					return fmt.Errorf("failed to get on_so value from report file %s: %w", us.report, err)
				}
				onBO, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onBOCol, rowWsReport+valueLocation))
				if err != nil {
					return fmt.Errorf("failed to get on_bo value from report file %s: %w", us.report, err)
				}

				// Convert onHand and onPO to ints
				var onHandInt, onPOInt int
				_, err = fmt.Sscan(onHand, &onHandInt)
				if err != nil {
					return fmt.Errorf("failed to convert on_hand value to int: %w", err)
				}
				_, err = fmt.Sscan(onPO, &onPOInt)
				if err != nil {
					return fmt.Errorf("failed to convert on_po value to int: %w", err)
				}

				// Add onSO and onBO together after converting them to ints
				var onSOInt, onBOInt int
				_, err = fmt.Sscan(onSO, &onSOInt)
				if err != nil {
					return fmt.Errorf("failed to convert on_so value to int: %w", err)
				}
				_, err = fmt.Sscan(onBO, &onBOInt)
				if err != nil {
					return fmt.Errorf("failed to convert on_bo value to int: %w", err)
				}
				onSOBOInt := onSOInt + onBOInt

				// Update the hotsheet
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.onHandCol, rowWsHotsheet), onHandInt)
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.onPOCol, rowWsHotsheet), onPOInt)
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.onSOBOCol, rowWsHotsheet), onSOBOInt)

				logger.Printf("Match found for SKU: %s | on_hand: %d | on_po: %d | on_so_bo: %d\n", skuWsHotsheet, onHandInt, onPOInt, onSOBOInt)
				wsReportPointer = rowWsReport + 1
				bar.Play(int64(rowWsHotsheet))
				break // Move to the next row in wsHotsheet once a match is found

			}
		}
	}

	if err := wbHotsheet.Save(); err != nil {
		return fmt.Errorf("failed to save hotsheet file: %w", err)
	}

	bar.Finish()
	return nil
}
