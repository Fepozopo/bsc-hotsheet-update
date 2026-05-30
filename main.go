package main

import (
	helpers "github.com/Fepozopo/bsc-hotsheet-update/helpers"
	"github.com/Fepozopo/bsc-hotsheet-update/internal/gui"
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

	if err := gui.Run(); err != nil {
		logger.Error("failed to run GUI", "err", err)
		return
	}

	logger.Info("application exited")
}
