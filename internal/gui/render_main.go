package gui

import (
	"fmt"
	"image/color"
	"strings"

	"github.com/aarzilli/nucular"
)

func (s *AppState) renderMainForm(w *nucular.Window) {
	w.Row(30).Dynamic(1)
	w.Label("Create Unified Hotsheets from Reports", "CC")

	w.Row(18).Dynamic(1)
	w.LabelColored("Inventory report is required. PO report and output directory are optional.", "CC", color.RGBA{R: 95, G: 95, B: 95, A: 255})

	s.renderSpacer(w, 6)
	s.renderPathField(w, "Inventory Report:", "Path to inventory report (.xlsx)", &s.inventoryEditor, s.browseInventory)
	s.renderSpacer(w, 6)
	s.renderPathField(w, "PO Report (optional):", "Path to PO report (.xlsx)", &s.poEditor, s.browsePO)
	s.renderSpacer(w, 6)
	s.renderPathField(w, "Output Directory (optional):", "Directory for generated files", &s.outputEditor, s.browseOutputDir)
	s.renderSpacer(w, 8)
	s.renderStatusLine(w)
	s.renderSpacer(w, 8)
	s.renderMainButtons(w)
}

func (s *AppState) renderPathField(w *nucular.Window, labelText, hintText string, editor *nucular.TextEditor, browseFn func()) {
	editor.Flags = s.pathEditorFlags()

	w.Row(20).Dynamic(1)
	w.Label(labelText, "LC")

	w.Row(28).Static(0, 90)
	editor.Edit(w)
	if w.ButtonText("Browse") && !s.isBusy() {
		browseFn()
	}

	w.Row(16).Dynamic(1)
	if strings.TrimSpace(string(editor.Buffer)) == "" {
		w.LabelColored(hintText, "LC", color.RGBA{R: 120, G: 120, B: 120, A: 255})
	} else {
		w.Label("", "LC")
	}
}

func (s *AppState) renderStatusLine(w *nucular.Window) {
	message := "Ready to generate hotsheets."
	statusColor := color.RGBA{R: 70, G: 110, B: 170, A: 255}

	switch {
	case s.updateInProgress:
		message = "Updating application..."
		statusColor = color.RGBA{R: 180, G: 120, B: 0, A: 255}
	case s.generateInProgress:
		message = "Generating hotsheets..."
		statusColor = color.RGBA{R: 180, G: 120, B: 0, A: 255}
	case s.updateCheckInProgress:
		message = "Checking for updates..."
		statusColor = color.RGBA{R: 70, G: 110, B: 170, A: 255}
	case strings.TrimSpace(s.updateStatusMessage) != "":
		message = s.updateStatusMessage
		statusColor = color.RGBA{R: 170, G: 105, B: 20, A: 255}
	case s.updateAvailable:
		message = fmt.Sprintf("Update available: %s (optional).", s.latestVersion)
		statusColor = color.RGBA{R: 30, G: 135, B: 70, A: 255}
	}

	w.Row(18).Dynamic(1)
	w.LabelColored(message, "LC", statusColor)
}

func (s *AppState) renderMainButtons(w *nucular.Window) {
	w.Row(34).Static(0, 150, 170, 80)
	w.Spacing(1)
	if w.ButtonText("Check for Updates") && !s.isBusy() && !s.updateCheckInProgress {
		s.startUpdateCheck(true)
	}
	if w.ButtonText("Generate Hotsheets") && !s.isBusy() && !s.updateCheckInProgress {
		s.startGenerate()
	}
	if w.ButtonText("Quit") {
		s.quit()
	}
}

func (s *AppState) renderSpacer(w *nucular.Window, height int) {
	w.Row(height).Dynamic(1)
	w.Label("", "LC")
}
