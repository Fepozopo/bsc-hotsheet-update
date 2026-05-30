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
	updateRepo           = "Fepozopo/bsc-hotsheet-update"
	doubleClickThreshold = 500 * time.Millisecond
	quitForceExitDelay   = 750 * time.Millisecond
)

func (s *AppState) isBusy() bool {
	return s.generateInProgress || s.updateInProgress
}

func (s *AppState) pathEditorFlags() nucular.EditFlags {
	flags := nucular.EditField
	if s.isBusy() {
		flags |= nucular.EditReadOnly
	}
	return flags
}

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
		outputs, err := hotsheet.CreateHotsheet(inv, po, outdir)
		s.queueEvent(generateCompletedEvent{Outputs: outputs, Err: err})
	}(inventoryPath, poPath, outputDir)
}

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

func (s *AppState) resetInputs() {
	setEditorText(&s.inventoryEditor, "")
	setEditorText(&s.poEditor, "")
	setEditorText(&s.outputEditor, "")
	s.selectedOutput = -1
	s.lastClickedOutput = -1
	s.lastClickAt = time.Time{}
	s.requestRedraw()
}

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

func (s *AppState) openSelectedOutput() {
	if s.selectedOutput >= 0 && s.selectedOutput < len(s.outputs) {
		OpenPath(s.outputs[s.selectedOutput])
	}
}

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

func (s *AppState) startUpdateCheck(showNoUpdates bool) {
	go func() {
		result, err := appupdate.CheckForUpdates(updateRepo, version.Version)
		s.queueEvent(updateCheckCompletedEvent{Result: result, Err: err, ShowNoUpdates: showNoUpdates})
	}()
}

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

func (s *AppState) handleUpdateCheckResult(result appupdate.CheckResult, err error, showNoUpdates bool) {
	if err != nil {
		s.updateAvailable = false
		s.latestVersion = ""
		s.latestAssetURL = ""
		s.updateStatusMessage = "Could not check for updates. Continuing without an update check."
		s.requestRedraw()
		return
	}

	s.updateStatusMessage = ""
	if !result.UpdateAvailable {
		s.updateAvailable = false
		if showNoUpdates {
			s.openInfoPopup("No Updates", "You are already running the latest version.")
		}
		return
	}

	if result.LatestVersion.String() == version.Version {
		s.updateAvailable = false
		return
	}

	s.updateAvailable = true
	s.latestVersion = result.LatestVersion.String()
	s.latestAssetURL = result.AssetURL
	s.openUpdateAvailablePopup()
}

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

func (s *AppState) quit() {
	if s.closingRequested {
		return
	}
	s.closingRequested = true

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

func (s *AppState) requestRedraw() {
	if s.mw != nil {
		s.mw.Changed()
	}
}

func editorText(editor *nucular.TextEditor) string {
	return strings.TrimSpace(string(editor.Buffer))
}

func setEditorText(editor *nucular.TextEditor, value string) {
	editor.Buffer = []rune(value)
	editor.Cursor = len(editor.Buffer)
	editor.SelectStart = 0
	editor.SelectEnd = 0
}
