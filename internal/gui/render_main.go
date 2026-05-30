package gui

import (
	"fmt"
	"image/color"
	"strings"

	"github.com/aarzilli/nucular"
)

// renderMainForm draws the main application window contents.
//
// The layout is intentionally kept close to the original Fyne-based UI: a title,
// three labeled path pickers, a status line, and the bottom action row.
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

// renderPathField draws a single labeled path editor with its Browse button and
// hint text.
func (s *AppState) renderPathField(w *nucular.Window, labelText, hintText string, editor *nucular.TextEditor, browseFn func()) {
	editor.Flags = s.pathEditorFlags()

	w.Row(20).Dynamic(1)
	w.Label(labelText, "LC")

	w.Row(28).Static(0, 90)
	editor.Edit(w)
	if w.ButtonText("Browse") && !s.isBusy() {
		browseFn()
	}

	// Nucular does not provide native placeholder support for TextEditor, so the
	// hint is rendered as a separate muted line below the field when it is empty.
	w.Row(16).Dynamic(1)
	if strings.TrimSpace(string(editor.Buffer)) == "" {
		w.LabelColored(hintText, "LC", color.RGBA{R: 120, G: 120, B: 120, A: 255})
	} else {
		w.Label("", "LC")
	}
}

// renderStatusLine shows a concise status message below the inputs.
//
// The message communicates whichever background action or update state is most
// relevant right now.
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

// renderMainButtons draws the primary action row at the bottom of the form.
//
// The requested layout keeps Quit and Check for Updates grouped on the left and
// Generate Hotsheets aligned on the right.
func (s *AppState) renderMainButtons(w *nucular.Window) {
	w.Row(34).Static(80, 14, 150, 0, 170)
	if w.ButtonText("Quit") {
		s.quit()
	}
	w.Label("", "LC")
	if w.ButtonText("Check for Updates") && !s.isBusy() && !s.updateCheckInProgress {
		s.startUpdateCheck(true)
	}
	w.Label("", "LC")
	if w.ButtonText("Generate Hotsheets") && !s.isBusy() && !s.updateCheckInProgress {
		s.startGenerate()
	}
}

// renderSpacer inserts a fixed-height blank row into the layout.
func (s *AppState) renderSpacer(w *nucular.Window, height int) {
	w.Row(height).Dynamic(1)
	w.Label("", "LC")
}
