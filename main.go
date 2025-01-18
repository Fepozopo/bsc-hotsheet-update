package main

import (
	"fmt"
	"log"
	"os"
	"time"

	"github.com/sqweek/dialog"
)

func main() {
	os.Exit(mainExitCode())
}

func mainExitCode() int {
	startTime := time.Now()

	// Get the current date
	currentDate := time.Now().Format("2006-01-02 15:04:05.000000000")

	// Create a new file path
	logFilePath := fmt.Sprintf("./logs/mainExitCode_%v.log", currentDate)

	// Create or open the log file
	logFile, err := os.OpenFile(logFilePath, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	if err != nil {
		fmt.Printf("error creating or opening log file: %v", err)
		return 1
	}
	defer logFile.Close()

	// Set the log output to the log file
	log.SetOutput(logFile)

	// Create a logger that writes to the log file
	logger := log.New(logFile, "INFO: ", log.Ldate|log.Ltime|log.Lshortfile)

	for {
		var hotsheet string
		fmt.Print("Which hotsheet do you want to update? (smd, bsc, 21c, exit): ")
		fmt.Scanln(&hotsheet)

		switch hotsheet {
		case "smd":
			fileSMDHotsheet, err := dialog.File().Title("Select the SMD HOTSHEET...").Filter("Excel Files", "*.xlsx").Load()
			if err != nil {
				logger.Printf("failed to open hotsheet file: %v", err)
				return 2
			}
			fileSMDStockReport, err := dialog.File().Title("Select the SMD Stock Report...").Filter("Excel Files", "*.xlsx").Load()
			if err != nil {
				logger.Printf("failed to open stock report file: %v", err)
				return 3
			}
			fileSMDSalesReport, err := dialog.File().Title("Select the SMD Sales Report...").Filter("Excel Files", "*.xlsx").Load()
			if err != nil {
				logger.Printf("failed to open sales report file: %v", err)
				return 4
			}
			fileSMDHotsheetNew, err := handlerCopyHotsheet("SMD", fileSMDHotsheet)
			if err != nil {
				logger.Printf("failed to copy hotsheet file: %v", err)
				return 5
			}

			// HOTSHEET | SECTION | REPORT | SKU | ON HAND | ON PO | ON SO | ON BO
			smdStock := UpdateStock{fileSMDHotsheetNew, "EVERYDAY", fileSMDStockReport, "E", "F", "I", "K"}
			smdStockHoliday := UpdateStock{fileSMDHotsheetNew, "HOLIDAY", fileSMDStockReport, "C", "D", "F", "H"}
			// HOTSHEET | SECTION | REPORT | SKU | YTD
			smdSales := UpdateSales{fileSMDHotsheetNew, "EVERYDAY", fileSMDSalesReport, "E", "Q"}
			smdSalesHoliday := UpdateSales{fileSMDHotsheetNew, "HOLIDAY", fileSMDSalesReport, "C", "O"}

			// Get user input for which sections to update
			for {
				var section string
				fmt.Print("Which section do you want to update? (everyday, holiday, all, exit): ")
				fmt.Scanln(&section)

				switch section {
				case "everyday":
					smdStock.handlerUpdateStock()
					fmt.Println("Stock updated successfully.")
					smdSales.handlerUpdateSales()
					fmt.Println("Sales updated successfully.")
				case "holiday":
					smdStockHoliday.handlerUpdateStock()
					fmt.Println("Holiday stock updated successfully.")
					smdSalesHoliday.handlerUpdateSales()
					fmt.Println("Holiday sales updated successfully.")
				case "all":
					smdStock.handlerUpdateStock()
					fmt.Println("Stock updated successfully.")
					smdStockHoliday.handlerUpdateStock()
					fmt.Println("Holiday stock updated successfully.")
					smdSales.handlerUpdateSales()
					fmt.Println("Sales updated successfully.")
					smdSalesHoliday.handlerUpdateSales()
					fmt.Println("Holiday sales updated successfully.")
				case "exit":
					printElapsedTime(startTime)
					return 0
				default:
					fmt.Println("Invalid input. Please enter 'everyday', 'holiday', 'all', or 'exit' (case sensitive).")
				}
			}

		case "bsc":
			// Similar logic for BSC hotsheet
			// TODO Implement file selection and updates for BSC here

		case "21c":
			// Similar logic for 21c hotsheet
			// TODO Implement file selection and updates for 21c here

		case "exit":
			printElapsedTime(startTime)
			return 0
		default:
			fmt.Println("Invalid input. Please enter 'smd', 'bsc', '21c', or 'exit' (case sensitive).")
		}
	}
}

func printElapsedTime(startTime time.Time) {
	timeElapsed := time.Since(startTime)
	fmt.Printf("Done!\nElapsed time: %v\n", timeElapsed)
}
