package main

import (
	"fmt"
	"log"
	"time"

	"github.com/sqweek/dialog"
)

func main() {
	startTime := time.Now()

	for {
		var hotsheet string
		fmt.Print("Which hotsheet do you want to update? (smd, bsc, 21c, exit): ")
		fmt.Scanln(&hotsheet)

		switch hotsheet {
		case "smd":
			fileSMDHotsheet, err := dialog.File().Title("Select the SMD HOTSHEET...").Filter("Excel Files", "*.xlsx").Load()
			if err != nil {
				log.Fatal(err)
			}
			fileSMDStockReport, err := dialog.File().Title("Select the SMD Stock Report...").Filter("Excel Files", "*.xlsx").Load()
			if err != nil {
				log.Fatal(err)
			}
			fileSMDSalesReport, err := dialog.File().Title("Select the SMD Sales Report...").Filter("Excel Files", "*.xlsx").Load()
			if err != nil {
				log.Fatal(err)
			}

			// HOTSHEET | SECTION | REPORT | SKU | ON HAND | ON PO | ON SO | ON BO
			smdStock := UpdateStock{fileSMDHotsheet, "EVERYDAY", fileSMDStockReport, "E", "F", "I", "K", "L"}
			smdStockHoliday := UpdateStock{fileSMDHotsheet, "HOLIDAY", fileSMDStockReport, "C", "D", "F", "H", "I"}
			// HOTSHEET | SECTION | REPORT | SKU | YTD
			smdSales := UpdateSales{fileSMDHotsheet, "EVERYDAY", fileSMDSalesReport, "E", "Q"}
			smdSalesHoliday := UpdateSales{fileSMDHotsheet, "HOLIDAY", fileSMDSalesReport, "C", "O"}

			// Get user input for which sections to update
			for {
				var section string
				fmt.Print("Which section do you want to update? (everyday, holiday, all, exit): ")
				fmt.Scanln(&section)

				switch section {
				case "everyday":
					smdStock.handlerUpdateStock()
					smdSales.handlerUpdateSales()
				case "holiday":
					smdStockHoliday.handlerUpdateStock()
					smdSalesHoliday.handlerUpdateSales()
				case "all":
					smdStock.handlerUpdateStock()
					smdStockHoliday.handlerUpdateStock()
					smdSales.handlerUpdateSales()
					smdSalesHoliday.handlerUpdateSales()
				case "exit":
				default:
					fmt.Println("Invalid input. Please enter 'everyday', 'holiday', 'all', or 'exit'.")
				}
			}

		case "bsc":
			// Similar logic for BSC hotsheet
			// TODO Implement file selection and updates for BSC here

		case "21c":
			// Similar logic for 21c hotsheet
			// TODO Implement file selection and updates for 21c here

		case "exit":
			return
		default:
			fmt.Println("Invalid input. Please enter 'smd', 'bsc', '21c', or 'exit'.")
		}

		timeElapsed := time.Since(startTime)
		fmt.Printf("Done!\nElapsed time: %s\n", timeElapsed)
	}
}
