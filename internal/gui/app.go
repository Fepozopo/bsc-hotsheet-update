package gui

import (
	"image"
	"os"
	"time"

	"github.com/aarzilli/nucular"
	nstyle "github.com/aarzilli/nucular/style"
)

const (
	// defaultWindowWidth and defaultWindowHeight define the initial size of the
	// top-level window before the user resizes it.
	defaultWindowWidth  = 900
	defaultWindowHeight = 600

	// defaultUIScale increases the effective size of the Nucular widgets so the
	// UI is comfortably readable on modern high-DPI displays, especially macOS.
	defaultUIScale = 2.0
)

// Run constructs the Nucular master window, binds it to the application state,
// and starts the GUI event loop.
//
// The function returns only if window creation fails before the Nucular main
// loop starts. Under normal operation the process exits when the window closes.
func Run() error {
	state := NewAppState()
	mw := nucular.NewMasterWindowSize(0, "Hotsheet Generator", image.Point{X: defaultWindowWidth, Y: defaultWindowHeight}, state.Update)

	// Nucular's shiny backend can leave the process alive after the native window
	// disappears, so a separate watchdog terminates the process once the window
	// is actually reported closed.
	mw.OnClose(func() {})
	mw.SetStyle(nstyle.FromTheme(nstyle.DefaultTheme, defaultUIScale))
	state.BindMasterWindow(mw)
	go watchForClosedWindow(mw)
	mw.Main()
	return nil
}

// watchForClosedWindow polls the master window and exits the process once the
// backend reports that the window is closed.
//
// This exists as a defensive workaround for backend-specific shutdown behavior,
// particularly on macOS with the shiny backend, where the window can close
// visually without the Go process terminating immediately.
func watchForClosedWindow(mw nucular.MasterWindow) {
	ticker := time.NewTicker(100 * time.Millisecond)
	defer ticker.Stop()

	for range ticker.C {
		if mw != nil && mw.Closed() {
			os.Exit(0)
		}
	}
}
