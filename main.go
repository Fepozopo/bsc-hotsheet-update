package main

import (
	"fmt"
	"time"

	"fyne.io/fyne/v2/app"
	helpers "github.com/Fepozopo/bsc-hotsheet-update/helpers"
)

// main is the entry point of the application. It initializes the logger, creates a new
// Fyne application, and prompts the user to select the product line and the paths to the
// hotsheet, stock report, and sales report files. If the selection is successful, it copies
// the hotsheet file and updates it based on the selected product line by invoking the
// appropriate helper functions. The function logs errors and exits early if any step fails.
func main() {
	startTime := time.Now()

	logger, logFile, err := helpers.CreateLogger("main", "", "", "ERROR")
	if err != nil {
		logger.Printf("failed to create log file: %v", err)
		return
	}
	defer logFile.Close()

	myApp := app.New()
	defer myApp.Quit()

	product, fileHotsheet, fileStockReport, fileSalesReport := selectFiles(myApp)

	// If no files are selected, exit
	if product == "" || fileHotsheet == "" || fileStockReport == "" || fileSalesReport == "" {
		logger.Printf("not all files were selected")
		return
	}

	// Copy the hotsheet
	fileHotsheetNew, err := helpers.CopyHotsheet(product, fileHotsheet)
	if err != nil {
		logger.Printf("failed to copy hotsheet file: %v", err)
		return
	}

	// Update the hotsheet
	var updateErr error
	switch product {
	case "smd":
		updateErr = helpers.CaseSMD(fileHotsheetNew, fileStockReport, fileSalesReport)
	case "bsc":
		updateErr = helpers.CaseBSC(fileHotsheetNew, fileStockReport, fileSalesReport)
	case "21c":
		updateErr = helpers.Case21C(fileHotsheetNew, fileStockReport, fileSalesReport)
	default:
		logger.Printf("unknown product: %s", product)
		return
	}
	if updateErr != nil {
		logger.Printf("failed to update %s hotsheet: %v", product, updateErr)
		return
	}

	fmt.Printf("Done!\nElapsed time: %v\n", time.Since(startTime))

	for i := 0; i < 5; i++ {
		fmt.Printf("Quitting in %d seconds...\n", 5-i)
		time.Sleep(1 * time.Second)
	}

}
