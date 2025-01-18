package main

import (
	"fmt"
	"log"
	"os"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type UpdateStock struct {
	hotsheet  string
	section   string
	report    string
	skuCol    string
	onHandCol string
	onPOCol   string
	onSOBOCol string
}

func (us *UpdateStock) handlerUpdateStock() error {
	// Get the current date
	currentDate := time.Now().Format("2006-01-02 15:04:05.000000000")

	// Create a new file path
	logFilePath := fmt.Sprintf("./logs/handlerUpdateStock_%v.log", currentDate)

	// Create or open the log file
	logFile, err := os.OpenFile(logFilePath, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	if err != nil {
		return fmt.Errorf("error creating or opening log file: %w", err)
	}
	defer logFile.Close()

	// Set the log output to the log file
	log.SetOutput(logFile)

	// Create a logger that writes to the log file
	logger := log.New(logFile, "INFO: ", log.Ldate|log.Ltime|log.Lshortfile)

	// Open the report workbook
	wbReport, err := excelize.OpenFile(us.report)
	if err != nil {
		log.Printf("failed to open report file %s: %v", us.report, err)
		return fmt.Errorf("failed to open report file %s: %w", us.report, err)
	}
	defer wbReport.Close()

	// Open the hotsheet workbook
	wbHotsheet, err := excelize.OpenFile(us.hotsheet)
	if err != nil {
		log.Printf("failed to open hotsheet file %s: %v", us.hotsheet, err)
		return fmt.Errorf("failed to open hotsheet file %s: %w", us.hotsheet, err)
	}
	defer wbHotsheet.Close()

	// Get the sheets
	wsReport := "Sheet1"
	wsHotsheet := us.section

	// Get the rows
	rowsHotsheet, err := wbHotsheet.GetRows(wsHotsheet)
	if err != nil {
		log.Printf("failed to get rows from hotsheet file %s: %v", us.hotsheet, err)
		return fmt.Errorf("failed to get rows from hotsheet file %s: %w", us.hotsheet, err)
	}
	rowsReport, err := wbReport.GetRows(wsReport)
	if err != nil {
		log.Printf("failed to get rows from report file %s: %v", us.report, err)
		return fmt.Errorf("failed to get rows from report file %s: %w", us.report, err)
	}

	skuCol := "A"        // 'A' column index in wsReport
	onHandCol := "K"     // 'K' column index in wsReport
	onPOCol := "M"       // 'M' column index in wsReport
	onSOCol := "P"       // 'P' column index in wsReport
	onBOCol := "T"       // 'T' column index in wsReport
	wsReportPointer := 1 // Start pointer for wsReport

	for rowWsHotsheet := 2; rowWsHotsheet < len(rowsHotsheet)+1; rowWsHotsheet++ {
		skuWsHotsheet, err := wbHotsheet.GetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.skuCol, rowWsHotsheet)) // SKU column in wsHotsheet
		if err != nil {
			log.Printf("failed to get SKU from hotsheet file %s: %v", us.hotsheet, err)
			return fmt.Errorf("failed to get SKU from hotsheet file %s: %w", us.hotsheet, err)
		}

		if skuWsHotsheet == "" {
			continue // Skip rows with no SKU in wsHotsheet
		}

		for rowWsReport := wsReportPointer; rowWsReport < len(rowsReport)+1; rowWsReport++ {
			skuWsReport, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", skuCol, rowWsReport)) // SKU in column 'A' in wsReport
			if err != nil {
				log.Printf("failed to get SKU from report file %s: %v", us.report, err)
				return fmt.Errorf("failed to get SKU from report file %s: %w", us.report, err)
			}

			if skuWsReport == "" {
				continue
			}

			// Skip the first sku that contains "MPN48"
			if skuWsReport == "MPN48 BOX  HUMMINGBIRD BOX OF 10 NOTECARD" {
				continue
			}

			logger.Printf("Comparing Hotsheet SKU: '%s' with Report SKU: '%s'\n", skuWsHotsheet, skuWsReport)
			if strings.Contains(skuWsReport, skuWsHotsheet) {
				testValueLocation, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onHandCol, rowWsReport+2))
				if err != nil {
					log.Printf("failed to get test value from report file %s: %v", us.report, err)
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
					log.Printf("failed to get on_hand value from report file %s: %v", us.report, err)
					return fmt.Errorf("failed to get on_hand value from report file %s: %w", us.report, err)
				}
				onPO, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onPOCol, rowWsReport+valueLocation))
				if err != nil {
					log.Printf("failed to get on_po value from report file %s: %v", us.report, err)
					return fmt.Errorf("failed to get on_po value from report file %s: %w", us.report, err)
				}
				onSO, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onSOCol, rowWsReport+valueLocation))
				if err != nil {
					log.Printf("failed to get on_so value from report file %s: %v", us.report, err)
					return fmt.Errorf("failed to get on_so value from report file %s: %w", us.report, err)
				}
				onBO, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onBOCol, rowWsReport+valueLocation))
				if err != nil {
					log.Printf("failed to get on_bo value from report file %s: %v", us.report, err)
					return fmt.Errorf("failed to get on_bo value from report file %s: %w", us.report, err)
				}

				// Add onSO and onBO together after converting them to ints
				var onSOInt int
				var onBOInt int
				_, err = fmt.Sscan(onSO, &onSOInt)
				if err != nil {
					log.Printf("failed to convert on_so value to int: %v", err)
					return fmt.Errorf("failed to convert on_so value to int: %w", err)
				}
				_, err = fmt.Sscan(onBO, &onBOInt)
				if err != nil {
					log.Printf("failed to convert on_bo value to int: %v", err)
					return fmt.Errorf("failed to convert on_bo value to int: %w", err)
				}
				onSOBO := onSOInt + onBOInt

				// Update the hotsheet
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.onHandCol, rowWsHotsheet), onHand)
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.onPOCol, rowWsHotsheet), onPO)
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.onSOBOCol, rowWsHotsheet), onSOBO)

				logger.Println("ROW | SKU | ON HAND | ON PO | ON SO/BO")
				logger.Printf("%v | %v | %v | %v | %v\n", rowWsHotsheet, skuWsHotsheet, onHand, onPO, onSOBO)
				wsReportPointer = rowWsReport + 1
				break // Move to the next row in wsHotsheet once a match is found

			}
		}
	}

	logger.Println("Saving file...")
	if err := wbHotsheet.Save(); err != nil {
		log.Printf("failed to save hotsheet file: %v", err)
		return fmt.Errorf("failed to save hotsheet file: %w", err)
	}

	return nil
}
