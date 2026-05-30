package gui

import (
	"time"

	"github.com/aarzilli/nucular"
	"github.com/aarzilli/nucular/rect"
)

// AppState stores the persistent UI state required by the immediate-mode Nucular GUI.
type AppState struct {
	mw                    nucular.MasterWindow
	events                chan UIEvent
	updateCheckStarted    bool
	updateCheckInProgress bool
	windowBounds          rect.Rect

	inventoryEditor nucular.TextEditor
	poEditor        nucular.TextEditor
	outputEditor    nucular.TextEditor

	outputs           []string
	selectedOutput    int
	lastClickedOutput int
	lastClickAt       time.Time

	generateInProgress  bool
	updateInProgress    bool
	closingRequested    bool
	updateAvailable     bool
	updateStatusMessage string
	latestVersion       string
	latestAssetURL      string
}

// NewAppState constructs the initial GUI state.
func NewAppState() *AppState {
	state := &AppState{
		events:            make(chan UIEvent, 16),
		selectedOutput:    -1,
		lastClickedOutput: -1,
		inventoryEditor:   newPathEditor(),
		poEditor:          newPathEditor(),
		outputEditor:      newPathEditor(),
	}
	return state
}

func (s *AppState) BindMasterWindow(mw nucular.MasterWindow) {
	s.mw = mw
}

// Update is the Nucular root update callback.
func (s *AppState) Update(w *nucular.Window) {
	s.windowBounds = w.Bounds
	s.drainEvents()
	if !s.updateCheckStarted {
		s.updateCheckStarted = true
		s.startUpdateCheck(false)
	}
	s.renderMainForm(w)
}

func (s *AppState) queueEvent(evt UIEvent) {
	s.events <- evt
	s.requestRedraw()
}

func (s *AppState) drainEvents() {
	for {
		select {
		case evt := <-s.events:
			switch e := evt.(type) {
			case generateCompletedEvent:
				s.handleGenerateResult(e.Outputs, e.Err)
			case updateCheckCompletedEvent:
				s.handleUpdateCheckResult(e.Result, e.Err, e.ShowNoUpdates)
			case selfUpdateCompletedEvent:
				s.handleSelfUpdateResult(e.Err)
			}
		default:
			return
		}
	}
}

func newPathEditor() nucular.TextEditor {
	return nucular.TextEditor{
		Flags:  nucular.EditField,
		Maxlen: 4096,
	}
}
