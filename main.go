package main

import (
	"fmt"
	"time"

	"fyne.io/fyne/v2/app"
	helpers "github.com/Fepozopo/bsc-hotsheet-update/helpers"
)

func main() {
	startTime := time.Now()

	logger, logFile, err := helpers.CreateLogger("main", "ERROR")
	if err != nil {
		logger.Printf("failed to create log file: %v", err)
		return
	}
	defer logFile.Close()

	myApp := app.New()
	defer myApp.Quit()

	// Start the main loop
	for {
		product, fileHotsheet, fileStockReport, fileSalesReport := selectFiles(myApp)

		// Copy the hotsheet
		fileHotsheetNew, err := helpers.CopyHotsheet(product, fileHotsheet)
		if err != nil {
			logger.Printf("failed to copy hotsheet file: %v", err)
			return
		}

		switch product {
		case "smd":
			err = helpers.CaseSMD(fileHotsheetNew, fileStockReport, fileSalesReport)
			if err != nil {
				logger.Printf("failed to update SMD hotsheet: %v", err)
				return
			}
			fmt.Printf("Done!\nElapsed time: %v\n", time.Since(startTime))
			return

		case "bsc":
			err = helpers.CaseBSC(fileHotsheetNew, fileStockReport, fileSalesReport)
			if err != nil {
				logger.Printf("failed to update BSC hotsheet: %v", err)
				return
			}
			fmt.Printf("Done!\nElapsed time: %v\n", time.Since(startTime))
			return

		case "21c":
			err = helpers.Case21C(fileHotsheetNew, fileStockReport, fileSalesReport)
			if err != nil {
				logger.Printf("failed to update 2021co hotsheet: %v", err)
				return
			}
			fmt.Printf("Done!\nElapsed time: %v\n", time.Since(startTime))
			return

		case "exit":
			return
		default:
			fmt.Println("Invalid input. Please enter 'smd', 'bsc', '21c', or 'exit' (case sensitive).")
		}
	}
}
