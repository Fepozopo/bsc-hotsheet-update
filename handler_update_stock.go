package main

import (
	"fmt"
	"strings"

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
	fmt.Println("ROW | SKU | ON HAND | ON PO | ON SO/BO")

	// Open the report workbook
	wbReport, err := excelize.OpenFile(us.report)
	if err != nil {
		return err
	}
	defer wbReport.Close()

	// Open the hotsheet workbook
	wbHotsheet, err := excelize.OpenFile(us.hotsheet)
	if err != nil {
		return err
	}
	defer wbHotsheet.Close()

	// Get the sheets
	wsReport := "Sheet1"
	wsHotsheet := us.section

	// Get the rows
	rowsHotsheet, err := wbHotsheet.GetRows(wsHotsheet)
	if err != nil {
		return err
	}
	rowsReport, err := wbReport.GetRows(wsReport)
	if err != nil {
		return err
	}

	skuCol := "A"        // 'A' column index in wsReport
	onHandCol := "K"     // 'K' column index in wsReport
	onPOCol := "M"       // 'M' column index in wsReport
	onSOCol := "P"       // 'P' column index in wsReport
	onBOCol := "T"       // 'T' column index in wsReport
	wsReportPointer := 1 // Start pointer for wsReport

	for rowWsHotsheet := 2; rowWsHotsheet < len(rowsHotsheet); rowWsHotsheet++ {
		skuWsHotsheet, err := wbHotsheet.GetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.skuCol, rowWsHotsheet)) // SKU column in wsHotsheet
		if err != nil {
			return err
		}

		if skuWsHotsheet == "" {
			continue // Skip rows with no SKU in wsHotsheet
		}

		for rowWsReport := wsReportPointer; rowWsReport < len(rowsReport); rowWsReport++ {
			skuWsReport, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", skuCol, rowWsReport)) // SKU in column 'A' in wsReport
			if err != nil {
				return err
			}

			if skuWsReport == "" {
				continue
			}

			if strings.Contains(skuWsReport, skuWsHotsheet) {
				var onHand, onPO, onSO, onBO string
				testReportValue, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onHandCol, rowWsReport+2))
				if err != nil {
					return err
				}

				// Skip the first sku that contains "MPN48"
				if skuWsReport == "MPN48 BOX  HUMMINGBIRD BOX OF 10 NOTECARD" {
					continue
				} else if testReportValue == "" {
					// Get the values (on_hand), (on_po), and (on_so, on_bo) in wsReport
					onHand, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onHandCol, rowWsReport+1))
					if err != nil {
						return err
					}
					onPO, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onPOCol, rowWsReport+1))
					if err != nil {
						return err
					}
					onSO, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onSOCol, rowWsReport+1))
					if err != nil {
						return err
					}
					onBO, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onBOCol, rowWsReport+1))
					if err != nil {
						return err
					}
				} else {
					// Get the values (on_hand), (on_po), and (on_so, on_bo) in wsReport
					onHand, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onHandCol, rowWsReport+2))
					if err != nil {
						return err
					}
					onPO, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onPOCol, rowWsReport+2))
					if err != nil {
						return err
					}
					onSO, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onSOCol, rowWsReport+2))
					if err != nil {
						return err
					}
					onBO, err = wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", onBOCol, rowWsReport+2))
					if err != nil {
						return err
					}
				}

				// Add onSO and onBO together after converting them to ints
				var onSOInt int
				var onBOInt int
				_, err = fmt.Sscan(onSO, &onSOInt)
				if err != nil {
					return err
				}
				_, err = fmt.Sscan(onBO, &onBOInt)
				if err != nil {
					return err
				}
				onSOBO := onSOInt + onBOInt
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.onHandCol, rowWsHotsheet), onHand)
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.onPOCol, rowWsHotsheet), onPO)
				wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.onSOBOCol, rowWsHotsheet), onSOBO)
				fmt.Printf("%v | %v | %v | %v | %v\n", rowWsHotsheet, skuWsHotsheet, onHand, onPO, onSOBO)

				wsReportPointer = rowWsReport + 1
				break // Move to the next row in wsHotsheet once a match is found

			}
		}
	}

	fmt.Println("Saving file...")
	if err := wbHotsheet.Save(); err != nil {
		return err
	}

	return nil
}
