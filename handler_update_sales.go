package main

import (
	"fmt"
	"log"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type UpdateSales struct {
	hotsheet string
	section  string
	report   string
	sku      int
	ytd      int
}

func (us *UpdateSales) handlerUpdateSales() {
	fmt.Println("ROW | SKU | YTD")

	// Open the report workbook
	wbReport, err := excelize.OpenFile(us.report)
	if err != nil {
		log.Fatal(err)
	}
	defer wbReport.Close()

	// Open the hotsheet workbook
	wbHotsheet, err := excelize.OpenFile(us.hotsheet)
	if err != nil {
		log.Fatal(err)
	}
	defer wbHotsheet.Close()

	// Get the sheets
	wsReport, err := wbReport.GetRows("Sheet1")
	if err != nil {
		log.Fatal(err)
	}

	wsHotsheet, err := wbHotsheet.GetRows(us.section)
	if err != nil {
		log.Fatal(err)
	}

	skuCol := 1          // 'A' column index in wsReport
	ytdCol := 19         // 'S' column index in wsReport
	kitCol := 10         // 'J' column index in wsReport
	wsReportPointer := 1 // Start pointer for wsReport

	for rowWsHotsheet := 2; rowWsHotsheet < len(wsHotsheet); rowWsHotsheet++ {
		skuWsHotsheet := wsHotsheet[rowWsHotsheet][us.sku]

		if skuWsHotsheet == "" {
			continue // Skip rows with no SKU in wsHotsheet
		}

		for rowWsReport := wsReportPointer; rowWsReport < len(wsReport); rowWsReport++ {
			skuWsReport := wsReport[rowWsReport][skuCol]

			if skuWsReport == "" {
				continue
			}

			if strings.TrimSpace(skuWsHotsheet) == strings.TrimSpace(skuWsReport) {
				if wsReport[rowWsReport+1][kitCol] == "Kit" {
					ytdValue := wsReport[rowWsReport+2][ytdCol]
					if strings.Contains(skuWsReport, "20-") || strings.Contains(skuWsReport, "21-") ||
						strings.Contains(skuWsReport, "22-") || strings.Contains(skuWsReport, "24-") ||
						strings.Contains(skuWsReport, "20F-") || strings.Contains(skuWsReport, "22F-") ||
						strings.Contains(skuWsReport, "24F-") {
						// Update (ytd) in wsHotsheet
						wsHotsheet[rowWsHotsheet][us.ytd] = ytdValue
						fmt.Println(rowWsHotsheet+1, "|", skuWsHotsheet, "|", ytdValue)
					} else {
						// Update (ytd) in wsHotsheet * 10
						ytdValueInt, err := strconv.Atoi(ytdValue)
						if err != nil {
							log.Fatal(err)
						}
						ytdValueString := strconv.Itoa(ytdValueInt * 10)
						wsHotsheet[rowWsHotsheet][us.ytd] = ytdValueString
						fmt.Println(rowWsHotsheet+1, "|", skuWsHotsheet, "|", ytdValue)
					}
				} else {
					// Update (ytd) in wsHotsheet
					wsHotsheet[rowWsHotsheet][us.ytd] = wsReport[rowWsReport+2][ytdCol]
					fmt.Println(rowWsHotsheet+1, "|", skuWsHotsheet, "|", wsReport[rowWsReport+2][ytdCol])
				}

				wsReportPointer = rowWsReport + 1
				break // Move to the next row in wsHotsheet once a match is found
			}
		}
	}

	fmt.Println("Saving file...")
	if err := wbHotsheet.SaveAs(us.hotsheet); err != nil {
		log.Fatal(err)
	}
}
