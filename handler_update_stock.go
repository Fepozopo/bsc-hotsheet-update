package main

import (
	"fmt"
	"log"
	"strings"

	"github.com/xuri/excelize/v2"
)

type UpdateStock struct {
	hotsheet string
	section  string
	report   string
	sku      int
	onHand   int
	onPO     int
	onSOBO   int
}

func (us *UpdateStock) handlerUpdateStock() {
	fmt.Println("ROW | SKU | ON HAND | ON PO | ON SO/BO")

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
	onHandCol := 11      // 'K' column index in wsReport
	onPOCol := 13        // 'M' column index in wsReport
	onSOCol := 16        // 'P' column index in wsReport
	onBOCol := 20        // 'T' column index in wsReport
	wsReportPointer := 1 // Start pointer for wsReport

	for rowWsHotsheet := 2; rowWsHotsheet < len(wsHotsheet); rowWsHotsheet++ {
		skuWsHotsheet := wsHotsheet[rowWsHotsheet][us.sku] // SKU column in wsHotsheet

		if skuWsHotsheet == "" {
			continue // Skip rows with no SKU in wsHotsheet
		}

		for rowWsReport := wsReportPointer; rowWsReport < len(wsReport); rowWsReport++ {
			skuWsReport := wsReport[rowWsReport][skuCol] // SKU in column 'A' in wsReport

			if skuWsReport == "" {
				continue
			}

			if strings.Contains(skuWsReport, skuWsHotsheet) {
				// Skip the first sku that contains "MPN48"
				if skuWsReport == "MPN48 BOX  HUMMINGBIRD BOX OF 10 NOTECARD" {
					continue
				}

				var onHand, onPO, onSOBO interface{}
				if wsReport[rowWsReport+2][onHandCol] == "" {
					// Update (on_hand), (on_po), and (on_so, on_bo) in wsHotsheet
					onHand = wsReport[rowWsReport+1][onHandCol]
					onPO = wsReport[rowWsReport+1][onPOCol]
					onSOBO = wsReport[rowWsReport+1][onSOCol] + wsReport[rowWsReport+1][onBOCol]
				} else {
					// Update (on_hand), (on_po), and (on_so, on_bo) in wsHotsheet
					onHand = wsReport[rowWsReport+2][onHandCol]
					onPO = wsReport[rowWsReport+2][onPOCol]
					onSOBO = wsReport[rowWsReport+2][onSOCol] + wsReport[rowWsReport+2][onBOCol]
				}

				// Type assertions
				onHandValue, _ := onHand.(string)
				onPOValue, _ := onPO.(string)
				onSOBOValue, _ := onSOBO.(string)

				// Set values in wsHotsheet
				wsHotsheet[rowWsHotsheet][us.onHand] = onHandValue
				wsHotsheet[rowWsHotsheet][us.onPO] = onPOValue
				wsHotsheet[rowWsHotsheet][us.onSOBO] = onSOBOValue

				fmt.Printf("%d | %s | %v | %v | %v\n", rowWsHotsheet, skuWsHotsheet, onHand, onPO, onSOBO)

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
