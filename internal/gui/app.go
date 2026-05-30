package gui

import (
	"image"
	"os"
	"time"

	"github.com/aarzilli/nucular"
	nstyle "github.com/aarzilli/nucular/style"
)

const (
	defaultWindowWidth  = 900
	defaultWindowHeight = 600
	defaultUIScale      = 2.0
)

// Run launches the Nucular application shell.
func Run() error {
	state := NewAppState()
	mw := nucular.NewMasterWindowSize(0, "Hotsheet Generator", image.Point{X: defaultWindowWidth, Y: defaultWindowHeight}, state.Update)
	mw.OnClose(func() {})
	mw.SetStyle(nstyle.FromTheme(nstyle.DefaultTheme, defaultUIScale))
	state.BindMasterWindow(mw)
	go watchForClosedWindow(mw)
	mw.Main()
	return nil
}

func watchForClosedWindow(mw nucular.MasterWindow) {
	ticker := time.NewTicker(100 * time.Millisecond)
	defer ticker.Stop()
	for range ticker.C {
		if mw != nil && mw.Closed() {
			os.Exit(0)
		}
	}
}
