package gui

import (
	"errors"
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/Fepozopo/bsc-hotsheet-update/hotsheet"
	appupdate "github.com/Fepozopo/bsc-hotsheet-update/internal/update"
	"github.com/Fepozopo/bsc-hotsheet-update/internal/version"
	"github.com/aarzilli/nucular"
)

const (
	// updateRepo is the GitHub repository identifier used by the built-in update
	// checker.
	updateRepo = "Fepozopo/bsc-hotsheet-update"

	// doubleClickThreshold defines how quickly two clicks on the same output row
	// must happen to be treated as a double-click.
	doubleClickThreshold = 500 * time.Millisecond

	// quitForceExitDelay is a safety timeout used when closing the application.
	// If the GUI backend does not terminate cleanly after a close request, the
	// process is exited explicitly.
	quitForceExitDelay = 750 * time.Millisecond
)

// isBusy reports whether the application is currently performing a long-running
// action that should temporarily disable conflicting UI actions.
func (s *AppState) isBusy() bool {
	return s.generateInProgress || s.updateInProgress
}

// pathEditorFlags returns the text editor flags appropriate for the three path
// fields in the main form.
//
// While generation or self-update is in progress the fields are switched to
// read-only to prevent the user from changing the underlying inputs mid-run.
func (s *AppState) pathEditorFlags() nucular.EditFlags {
	flags := nucular.EditField
	if s.isBusy() {
		flags |= nucular.EditReadOnly
	}
	return flags
}

// browseInventory opens the native file picker and stores the selected
// inventory report path in the inventory editor.
func (s *AppState) browseInventory() {
	path, err := pickFile()
	if err != nil {
		if errors.Is(err, errDialogCancelled) {
			return
		}
		s.openErrorPopup("Browse Error", err.Error())
		return
	}
	setEditorText(&s.inventoryEditor, path)
	s.requestRedraw()
}

// browsePO opens the native file picker and stores the selected PO report path
// in the optional PO editor.
func (s *AppState) browsePO() {
	path, err := pickFile()
	if err != nil {
		if errors.Is(err, errDialogCancelled) {
			return
		}
		s.openErrorPopup("Browse Error", err.Error())
		return
	}
	setEditorText(&s.poEditor, path)
	s.requestRedraw()
}

// browseOutputDir opens the native directory picker and stores the chosen
// output directory path.
func (s *AppState) browseOutputDir() {
	path, err := pickDirectory()
	if err != nil {
		if errors.Is(err, errDialogCancelled) {
			return
		}
		s.openErrorPopup("Browse Error", err.Error())
		return
	}
	setEditorText(&s.outputEditor, path)
	s.requestRedraw()
}

// startGenerate validates the required inputs, shows the progress popup, and
// launches hotsheet generation in a background goroutine.
func (s *AppState) startGenerate() {
	if s.isBusy() {
		return
	}

	inventoryPath := editorText(&s.inventoryEditor)
	if inventoryPath == "" {
		s.openErrorPopup("Missing Inventory Report", "Inventory report is required")
		return
	}

	poPath := editorText(&s.poEditor)
	outputDir := editorText(&s.outputEditor)

	s.generateInProgress = true
	s.openGenerateProgressPopup()
	s.requestRedraw()

	go func(inv, po, outdir string) {
		outputs, err := hotsheet.Generate(inv, po, outdir)
		s.queueEvent(generateCompletedEvent{Outputs: outputs, Err: err})
	}(inventoryPath, poPath, outputDir)
}

// handleGenerateResult applies the result of a completed generation run back to
// the UI state.
//
// Successful runs open the results popup; failed runs surface the error in a
// modal popup and leave the main form intact.
func (s *AppState) handleGenerateResult(outputs []string, err error) {
	s.generateInProgress = false
	if err != nil {
		s.openErrorPopup("Generation Failed", err.Error())
		s.requestRedraw()
		return
	}

	s.outputs = outputs
	s.selectedOutput = -1
	s.lastClickedOutput = -1
	s.lastClickAt = time.Time{}
	s.openOutputsPopup()
	s.requestRedraw()
}

// resetInputs clears the three main path fields and resets any result-list
// selection state so the user can start a fresh run.
func (s *AppState) resetInputs() {
	setEditorText(&s.inventoryEditor, "")
	setEditorText(&s.poEditor, "")
	setEditorText(&s.outputEditor, "")
	s.selectedOutput = -1
	s.lastClickedOutput = -1
	s.lastClickAt = time.Time{}
	s.requestRedraw()
}

// handleOutputClick updates selection state for the results list and opens the
// file when the same row is clicked twice inside doubleClickThreshold.
func (s *AppState) handleOutputClick(index int) {
	now := time.Now()
	if s.lastClickedOutput == index && now.Sub(s.lastClickAt) <= doubleClickThreshold {
		s.selectedOutput = index
		s.openSelectedOutput()
	}
	s.selectedOutput = index
	s.lastClickedOutput = index
	s.lastClickAt = now
	s.requestRedraw()
}

// openSelectedOutput opens the currently selected output file, if any.
func (s *AppState) openSelectedOutput() {
	if s.selectedOutput >= 0 && s.selectedOutput < len(s.outputs) {
		OpenPath(s.outputs[s.selectedOutput])
	}
}

// openSelectedOutputFolder opens the containing directory for the selected
// output file.
//
// If no row is selected, the function falls back to the first generated output,
// matching the behavior of the earlier Fyne implementation.
func (s *AppState) openSelectedOutputFolder() {
	if len(s.outputs) == 0 {
		return
	}
	if s.selectedOutput >= 0 && s.selectedOutput < len(s.outputs) {
		OpenFolderForFile(s.outputs[s.selectedOutput])
		return
	}
	OpenFolderForFile(s.outputs[0])
}

// startUpdateCheck launches an asynchronous update check against GitHub.
//
// The check runs on startup and can also be triggered manually. Re-entrant runs
// are ignored to avoid stacking duplicate requests and popups.
func (s *AppState) startUpdateCheck(showNoUpdates bool) {
	if s.updateCheckInProgress || s.updateInProgress || s.closingRequested {
		return
	}

	s.updateCheckInProgress = true
	s.updateStatusMessage = "Checking for updates..."
	s.requestRedraw()

	go func() {
		result, err := appupdate.CheckForUpdates(updateRepo, version.Version)
		s.queueEvent(updateCheckCompletedEvent{Result: result, Err: err, ShowNoUpdates: showNoUpdates})
	}()
}

// beginSelfUpdate starts downloading and applying the selected update in the
// background.
//
// The current version remains usable if the update fails; only a successful
// update path restarts and exits the running process.
func (s *AppState) beginSelfUpdate() {
	if s.isBusy() || strings.TrimSpace(s.latestAssetURL) == "" {
		return
	}

	s.updateInProgress = true
	s.openUpdateProgressPopup()
	s.requestRedraw()

	go func(assetURL string) {
		exe, err := os.Executable()
		if err != nil {
			s.queueEvent(selfUpdateCompletedEvent{Err: fmt.Errorf("could not locate executable: %w", err)})
			return
		}
		if err := appupdate.ApplyUpdate(assetURL, exe); err != nil {
			s.queueEvent(selfUpdateCompletedEvent{Err: fmt.Errorf("update failed: %w", err)})
			return
		}
		if err := appupdate.RestartExecutable(exe, os.Args[1:]); err != nil {
			s.queueEvent(selfUpdateCompletedEvent{Err: fmt.Errorf("failed to restart: %w", err)})
			return
		}
		s.queueEvent(selfUpdateCompletedEvent{})
	}(s.latestAssetURL)
}

// handleUpdateCheckResult merges the outcome of a completed update check into
// the application state and decides whether to show any popup.
//
// A failed check is intentionally non-fatal: the user can keep using the app
// even when GitHub is unreachable.
func (s *AppState) handleUpdateCheckResult(result appupdate.CheckResult, err error, showNoUpdates bool) {
	s.updateCheckInProgress = false
	if err != nil {
		s.updateAvailable = false
		s.latestVersion = ""
		s.latestAssetURL = ""
		s.updateStatusMessage = "Could not check for updates. Continuing without an update check."
		if showNoUpdates {
			s.openInfoPopup("Update Check Failed", "Could not check for updates right now. Continuing without an update check.")
		}
		s.requestRedraw()
		return
	}

	s.updateStatusMessage = ""
	if !result.UpdateAvailable {
		s.updateAvailable = false
		s.latestVersion = ""
		s.latestAssetURL = ""
		if showNoUpdates {
			s.openInfoPopup("No Updates", "You are already running the latest version.")
		}
		return
	}

	// The explicit string check is defensive; semver comparison should already
	// cover equality, but keeping this branch makes the intent obvious.
	if result.LatestVersion.String() == version.Version {
		s.updateAvailable = false
		return
	}

	s.updateAvailable = true
	s.latestVersion = result.LatestVersion.String()
	s.latestAssetURL = result.AssetURL
	s.openUpdateAvailablePopup()
}

// handleSelfUpdateResult applies the result of the background self-update.
func (s *AppState) handleSelfUpdateResult(err error) {
	s.updateInProgress = false
	if err != nil {
		s.updateStatusMessage = "Update failed. You can continue using the current version."
		s.openErrorPopup("Update Failed", err.Error())
		s.requestRedraw()
		return
	}
	s.quit()
}

// quit initiates application shutdown and guarantees process termination even
// if the GUI backend does not exit cleanly on its own.
func (s *AppState) quit() {
	if s.closingRequested {
		return
	}
	s.closingRequested = true

	// The fallback exit protects against backend shutdown deadlocks. The native
	// window is still asked to close first so the UI can terminate cleanly when
	// possible.
	go func() {
		time.Sleep(quitForceExitDelay)
		os.Exit(0)
	}()

	if s.mw == nil {
		return
	}

	go func(mw nucular.MasterWindow) {
		mw.Close()
	}(s.mw)
}

// requestRedraw schedules another Nucular frame if the master window exists.
func (s *AppState) requestRedraw() {
	if s.mw != nil {
		s.mw.Changed()
	}
}

// editorText returns the trimmed contents of a Nucular text editor.
func editorText(editor *nucular.TextEditor) string {
	return strings.TrimSpace(string(editor.Buffer))
}

// setEditorText replaces the entire contents of a Nucular text editor and moves
// the cursor to the end of the new text.
func setEditorText(editor *nucular.TextEditor, value string) {
	editor.Buffer = []rune(value)
	editor.Cursor = len(editor.Buffer)
	editor.SelectStart = 0
	editor.SelectEnd = 0
}
