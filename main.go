package main

import (
	"fyne.io/fyne/v2/app"
	helpers "github.com/Fepozopo/bsc-hotsheet-update/helpers"
)

func main() {
	logger, logFile, err := helpers.CreateLogger("main", "", "", "ERROR")
	if err != nil {
		// If we cannot create the logger, we cannot proceed reliably.
		// Fail early; the UI flow depends on this logging setup.
		return
	}
	defer logFile.Close()

	myApp := app.New()
	defer myApp.Quit()

	// The UI-driven flow in selectFiles handles generation and result display.
	selectFiles(myApp)

	logger.Println("application exited")
}
