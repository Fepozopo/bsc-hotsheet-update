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

	inventoryReport, poReport, outputDir := selectFiles(myApp)

	// If no reports are selected, exit
	if inventoryReport == "" || poReport == "" {
		logger.Printf("not all report files were selected")
		return
	}

	outputs, err := hotsheet.CreateFromReports(inventoryReport, poReport, outputDir)
	if err != nil {
		logger.Printf("failed to create hotsheets: %v", err)
		return
	}
	for _, out := range outputs {
		fmt.Printf("Created: %s\n", out)
	}

	for i := 0; i < 3; i++ {
		fmt.Printf("Quitting in %d seconds...\n", 3-i)
		time.Sleep(1 * time.Second)
	}

}
