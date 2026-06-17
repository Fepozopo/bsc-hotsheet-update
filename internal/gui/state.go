package gui

import (
	"runtime"
	"time"

	"github.com/aarzilli/nucular"
	"github.com/aarzilli/nucular/rect"
	"golang.org/x/mobile/event/key"
)

// AppState stores the persistent state that drives the immediate-mode GUI.
//
// Unlike retained-mode toolkits, Nucular does not keep long-lived widgets with
// their own internal application logic. Instead, the application owns the state
// explicitly and redraws the interface from that state on every frame.
type popupKind int

const (
	popupNone popupKind = iota
	popupMessage
	popupGenerateProgress
	popupUpdateAvailable
	popupUpdateProgress
	popupOutputs
)

// AppState contains all mutable state owned by the GUI layer.
//
// Nucular redraws the interface from application-owned data on every frame, so
// this struct is the single source of truth for window state, form inputs,
// background-job status, and transient popup state.
type AppState struct {
	mw                    nucular.MasterWindow
	events                chan UIEvent
	updateCheckStarted    bool
	updateCheckInProgress bool
	windowBounds          rect.Rect
	currentPopup          popupKind

	// Text editors back the three path fields in the main form.
	inventoryEditor nucular.TextEditor
	poEditor        nucular.TextEditor
	outputEditor    nucular.TextEditor

	// Output selection state is tracked separately from the rendered list because
	// the immediate-mode UI is rebuilt each frame.
	outputs                   []string
	selectedOutput            int
	selectedOutputNeedsScroll bool
	lastClickedOutput         int
	lastClickAt               time.Time

	generateInProgress bool
	// generateProgress and generateProgressMessage are written only from the UI
	// event-drain path. Background goroutines must send generateProgressEvent
	// values instead of mutating these fields directly.
	generateProgress        int
	generateProgressMessage string
	updateInProgress        bool
	closingRequested        bool
	updateAvailable         bool
	updateStatusMessage     string
	latestVersion           string
	latestAssetURL          string
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

// BindMasterWindow stores the top-level master window reference once it has
// been created.
func (s *AppState) BindMasterWindow(mw nucular.MasterWindow) {
	s.mw = mw
}

// Update is the root Nucular update callback.
//
// Each frame captures the latest root window bounds, drains any pending
// background events, kicks off the startup update check once, and then redraws
// the main form.
func (s *AppState) Update(w *nucular.Window) {
	s.windowBounds = w.Bounds
	s.drainEvents()
	if !s.updateCheckStarted {
		s.updateCheckStarted = true
		s.startUpdateCheck(false)
	}
	s.renderMainForm(w)
}

// queueEvent posts a background result back to the UI state and requests a new
// frame so the event can be applied promptly.
func (s *AppState) queueEvent(evt UIEvent) {
	s.events <- evt
	s.requestRedraw()
}

// drainEvents applies all pending background results to the current AppState.
func (s *AppState) drainEvents() {
	for {
		select {
		case evt := <-s.events:
			switch e := evt.(type) {
			case generateProgressEvent:
				s.handleGenerateProgress(e.Progress)
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

// newPathEditor returns a TextEditor configured for single-line path entry.
func newPathEditor() nucular.TextEditor {
	return nucular.TextEditor{
		Flags:  nucular.EditField,
		Maxlen: 4096,
	}
}

// shortcutModifier returns the platform-appropriate application shortcut
// modifier: Command on macOS, Control elsewhere.
func shortcutModifier() key.Modifiers {
	if runtime.GOOS == "darwin" {
		return key.ModMeta
	}
	return key.ModControl
}

// anyEditorActive reports whether one of the main form text inputs currently
// owns keyboard focus.
func (s *AppState) anyEditorActive() bool {
	return s.inventoryEditor.Active || s.poEditor.Active || s.outputEditor.Active
}
