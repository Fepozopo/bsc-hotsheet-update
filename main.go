package main

import (
	"fmt"
	"time"

	"fyne.io/fyne/v2/app"
	helpers "github.com/Fepozopo/bsc-hotsheet-update/helpers"
)

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

}
