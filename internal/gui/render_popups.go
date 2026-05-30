package gui

import (
	"fmt"

	"github.com/aarzilli/nucular"
	"github.com/aarzilli/nucular/rect"
)

func (s *AppState) openErrorPopup(title, message string) {
	s.openMessagePopup(title, message, "OK")
}

func (s *AppState) openInfoPopup(title, message string) {
	s.openMessagePopup(title, message, "OK")
}

func (s *AppState) openMessagePopup(title, message, dismissText string) {
	s.mw.PopupOpen(title, nucular.WindowMovable|nucular.WindowTitle|nucular.WindowDynamic|nucular.WindowNoScrollbar, rect.Rect{X: 60, Y: 60, W: 520, H: 190}, true, func(w *nucular.Window) {
		w.Row(95).Dynamic(1)
		w.LabelWrap(message)

		w.Row(32).Static(0, 90)
		w.Spacing(1)
		if w.ButtonText(dismissText) {
			w.Close()
		}
	})
}

func (s *AppState) openGenerateProgressPopup() {
	s.mw.PopupOpen("Generating Hotsheets", nucular.WindowMovable|nucular.WindowTitle|nucular.WindowDynamic|nucular.WindowNoScrollbar, rect.Rect{X: 90, Y: 80, W: 430, H: 180}, true, s.renderGenerateProgressPopup)
}

func (s *AppState) renderGenerateProgressPopup(w *nucular.Window) {
	if !s.generateInProgress {
		w.Close()
		return
	}

	w.Row(28).Dynamic(1)
	w.Label("Generating hotsheets...", "LC")
	w.Row(60).Dynamic(1)
	w.LabelWrap("Please wait while the reports are processed. Closing this popup will not cancel the generation.")
	w.Row(32).Static(0, 90)
	w.Spacing(1)
	if w.ButtonText("Cancel") {
		w.Close()
	}
}

func (s *AppState) openUpdateAvailablePopup() {
	s.mw.PopupOpen("Update Available", nucular.WindowMovable|nucular.WindowTitle|nucular.WindowDynamic|nucular.WindowNoScrollbar, rect.Rect{X: 55, Y: 55, W: 600, H: 230}, true, s.renderUpdateAvailablePopup)
}

func (s *AppState) renderUpdateAvailablePopup(w *nucular.Window) {
	if s.updateInProgress {
		w.Close()
		return
	}

	message := fmt.Sprintf("A new version (%s) is available. You can update now, or continue using the current version.", s.latestVersion)
	w.Row(118).Dynamic(1)
	w.LabelWrap(message)
	w.Row(32).Static(0, 110, 110)
	w.Spacing(1)
	if w.ButtonText("Update") {
		s.beginSelfUpdate()
		return
	}
	if w.ButtonText("Continue") {
		w.Close()
	}
}

func (s *AppState) openUpdateProgressPopup() {
	s.mw.PopupOpen("Updating", nucular.WindowMovable|nucular.WindowTitle|nucular.WindowDynamic|nucular.WindowNoScrollbar, rect.Rect{X: 90, Y: 80, W: 430, H: 180}, true, s.renderUpdateProgressPopup)
}

func (s *AppState) renderUpdateProgressPopup(w *nucular.Window) {
	if !s.updateInProgress {
		w.Close()
		return
	}

	w.Row(28).Dynamic(1)
	w.Label("Updating application...", "LC")
	w.Row(60).Dynamic(1)
	w.LabelWrap("Please wait while the new version is downloaded and applied. Closing this popup will not stop the update.")
	w.Row(32).Static(0, 90)
	w.Spacing(1)
	if w.ButtonText("Cancel") {
		w.Close()
	}
}

func (s *AppState) openOutputsPopup() {
	s.mw.PopupOpen("Created Hotsheets", nucular.WindowMovable|nucular.WindowTitle|nucular.WindowDynamic|nucular.WindowNoScrollbar, rect.Rect{X: 40, Y: 30, W: 620, H: 400}, true, s.renderOutputsPopup)
}

func (s *AppState) renderOutputsPopup(w *nucular.Window) {
	if len(s.outputs) == 0 {
		w.Row(20).Dynamic(1)
		w.Label("No files were created.", "LC")
		w.Row(220).Dynamic(1)
		w.LabelWrap("No output files were returned by the generator.")
	} else {
		w.Row(20).Dynamic(1)
		w.Label(fmt.Sprintf("Created files (%d):", len(s.outputs)), "LC")
		w.Row(18).Dynamic(1)
		w.Label("Double-click a file to open it.", "LC")
		w.Row(235).Dynamic(1)
		if gl, gw := nucular.GroupListStart(w, len(s.outputs), "created-hotsheets", nucular.WindowBorder|nucular.WindowNoHScrollbar); gw != nil {
			gl.SkipToVisible(22)
			gw.Row(22).Dynamic(1)
			for gl.Next() {
				idx := gl.Index()
				selected := idx == s.selectedOutput
				if gw.SelectableLabel(s.outputs[idx], "LC", &selected) {
					s.handleOutputClick(idx)
				}
			}
		}
	}

	w.Row(32).Static(0, 120, 80)
	w.Spacing(1)
	if w.ButtonText("Open Folder") {
		s.openSelectedOutputFolder()
	}
	if w.ButtonText("Done") {
		s.resetInputs()
		s.outputs = nil
		w.Close()
	}
}
