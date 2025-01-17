package main

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

type UpdateSales struct {
	hotsheet string
	section  string
	report   string
	skuCol   string
	ytdCol   string
}

func (us *UpdateSales) handlerUpdateSales() error {
	fmt.Println("ROW | SKU | YTD")

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
	ytdCol := "S"        // 'S' column index in wsReport
	kitCol := "J"        // 'J' column index in wsReport
	wsReportPointer := 1 // Start pointer for wsReport

	for rowWsHotsheet := 2; rowWsHotsheet < len(rowsHotsheet); rowWsHotsheet++ {
		skuWsHotsheet, err := wbHotsheet.GetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.skuCol, rowWsHotsheet))
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

			if strings.TrimSpace(skuWsHotsheet) == strings.TrimSpace(skuWsReport) {
				isKit, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", kitCol, rowWsReport+1))
				if err != nil {
					return err
				}
				ytdValue, err := wbReport.GetCellValue(wsReport, fmt.Sprintf("%s%d", ytdCol, rowWsReport+2))
				if err != nil {
					return err
				}
				if isKit == "Kit" {
					if strings.Contains(skuWsReport, "20-") || strings.Contains(skuWsReport, "21-") ||
						strings.Contains(skuWsReport, "22-") || strings.Contains(skuWsReport, "24-") ||
						strings.Contains(skuWsReport, "20F-") || strings.Contains(skuWsReport, "22F-") ||
						strings.Contains(skuWsReport, "24F-") {
						// Update (ytd) in wsHotsheet
						wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.ytdCol, rowWsHotsheet+2), ytdValue)
						fmt.Println(rowWsHotsheet, "|", skuWsHotsheet, "|", ytdValue)
					} else {
						// Convert to int in order to multiply by 10
						var ytdValueInt int
						_, err := fmt.Sscan(ytdValue, &ytdValueInt)
						if err != nil {
							return err
						}
						ytdValue10x := ytdValueInt * 10
						wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.ytdCol, rowWsHotsheet), ytdValue10x)
						fmt.Println(rowWsHotsheet, "|", skuWsHotsheet, "|", ytdValue)
					}
				} else {
					// Update (ytd) in wsHotsheet
					wbHotsheet.SetCellValue(wsHotsheet, fmt.Sprintf("%s%d", us.ytdCol, rowWsHotsheet+2), ytdValue)
					fmt.Println(rowWsHotsheet, "|", skuWsHotsheet, "|", ytdValue)
				}

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
