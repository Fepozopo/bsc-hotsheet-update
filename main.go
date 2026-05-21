package main

import (
	"fyne.io/fyne/v2/app"
	helpers "github.com/Fepozopo/bsc-hotsheet-update/helpers"
)

// main wires up the application logger and launches the GUI flow.
func main() {
	logger, logCloser, err := helpers.CreateSlogLogger("main", "DEBUG")
	if err != nil {
		// If we cannot create the logger, we cannot proceed reliably.
		// Fail early; the UI flow depends on this logging setup.
		return
	}
	defer func() {
		_ = logCloser.Close()
	}()

	myApp := app.New()
	defer myApp.Quit()

	// The UI-driven flow in selectFiles handles generation and result display.
	selectFiles(myApp)

	logger.Info("application exited")
}
