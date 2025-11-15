package main

import (
	"fmt"
	"time"

	"fyne.io/fyne/v2/app"
	helpers "github.com/Fepozopo/bsc-hotsheet-update/helpers"
	"github.com/Fepozopo/bsc-hotsheet-update/hotsheet"
)

// main is the entry point of the application. It initializes the logger, creates a new
// Fyne application, and prompts the user to select the product line and the paths to the
// hotsheet, inventory report, and PO report files. If the selection is successful, it copies
// the hotsheet file and updates it based on the selected product line by invoking the
// appropriate helper functions. The function logs errors and exits early if any step fails.
func main() {
	logger, logFile, err := helpers.CreateLogger("main", "", "", "ERROR")
	if err != nil {
		logger.Printf("failed to create log file: %v", err)
		return
	}
	defer logFile.Close()

	myApp := app.New()
	defer myApp.Quit()

	selection, hotsheetPaths, inventoryReport, poReport := selectFiles(myApp)

	// If no files are selected, exit
	if selection == "" || len(hotsheetPaths) == 0 || inventoryReport == "" || poReport == "" {
		logger.Printf("not all files were selected")
		return
	}

	if selection == "All" {
		products := []struct {
			name       string
			updateFunc func(string, string, string) error
		}{
			{"21c", hotsheet.Case21C},
			{"BSC", hotsheet.CaseBSC},
			{"BJP", hotsheet.CaseBJP},
			{"SMD", hotsheet.CaseSMD},
		}
		for i, p := range products {
			if i >= len(hotsheetPaths) {
				logger.Printf("missing hotsheet file for %s", p.name)
				continue
			}
			fileHotsheet := hotsheetPaths[i]
			if fileHotsheet == "" {
				logger.Printf("no hotsheet file selected for %s", p.name)
				continue
			}
			fileHotsheetNew, err := hotsheet.CopyHotsheet(p.name, fileHotsheet)
			if err != nil {
				logger.Printf("failed to copy %s hotsheet file: %v", p.name, err)
				continue
			}
			if err := p.updateFunc(fileHotsheetNew, inventoryReport, poReport); err != nil {
				logger.Printf("failed to update %s hotsheet: %v", p.name, err)
			} else {
				fmt.Printf("%s hotsheet updated successfully.\n", p.name)
			}
		}
	} else {
		fileHotsheet := hotsheetPaths[0]
		if fileHotsheet == "" {
			logger.Printf("no hotsheet file selected")
			return
		}
		// Copy the hotsheet
		fileHotsheetNew, err := hotsheet.CopyHotsheet(selection, fileHotsheet)
		if err != nil {
			logger.Printf("failed to copy hotsheet file: %v", err)
			return
		}

		// Update the hotsheet
		var updateErr error
		switch selection {
		case "21c":
			updateErr = hotsheet.Case21C(fileHotsheetNew, inventoryReport, poReport)
		case "BSC":
			updateErr = hotsheet.CaseBSC(fileHotsheetNew, inventoryReport, poReport)
		case "BJP":
			updateErr = hotsheet.CaseBJP(fileHotsheetNew, inventoryReport, poReport)
		case "SMD":
			updateErr = hotsheet.CaseSMD(fileHotsheetNew, inventoryReport, poReport)
		default:
			logger.Printf("unknown product: %s", selection)
			return
		}
		if updateErr != nil {
			logger.Printf("failed to update %s hotsheet: %v", selection, updateErr)
			return
		}
	}

	for i := 0; i < 3; i++ {
		fmt.Printf("Quitting in %d seconds...\n", 3-i)
		time.Sleep(1 * time.Second)
	}

}
