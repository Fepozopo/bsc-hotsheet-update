package gui

import (
	"fmt"
	"strings"
	"unicode/utf8"

	"github.com/aarzilli/nucular"
	"github.com/aarzilli/nucular/rect"
	"golang.org/x/mobile/event/key"
)

// openErrorPopup shows a modal error popup with a single dismiss button.
func (s *AppState) openErrorPopup(title, message string) {
	s.openMessagePopup(title, message, "OK")
}

// openInfoPopup shows a modal informational popup with a single dismiss button.
func (s *AppState) openInfoPopup(title, message string) {
	s.openMessagePopup(title, message, "OK")
}

// openMessagePopup renders a small centered popup for one-off informational or
// error messages.
func (s *AppState) openMessagePopup(title, message, dismissText string) {
	s.currentPopup = popupMessage
	s.mw.PopupOpen(title, nucular.WindowMovable|nucular.WindowTitle|nucular.WindowDynamic|nucular.WindowNoScrollbar, s.centeredPopupRect(560, 220), true, func(w *nucular.Window) {
		if s.handlePopupEscape(w) {
			return
		}
		s.renderPopupMessage(w, message, 48)
		w.Row(56).Dynamic(1)
		w.Label("", "LC")
		w.Row(32).Static(0, 110, 0)
		w.Label("", "LC")
		if w.ButtonText(dismissText) {
			s.closePopup(w)
		}
		w.Label("", "LC")
	})
}

// openGenerateProgressPopup shows the modal progress popup for hotsheet
// generation.
func (s *AppState) openGenerateProgressPopup() {
	s.currentPopup = popupGenerateProgress
	s.mw.PopupOpen("Generating Hotsheets", nucular.WindowMovable|nucular.WindowTitle|nucular.WindowDynamic|nucular.WindowNoScrollbar, s.centeredPopupRect(460, 200), true, s.renderGenerateProgressPopup)
}

// renderGenerateProgressPopup draws the content of the generation progress
// popup.
func (s *AppState) renderGenerateProgressPopup(w *nucular.Window) {
	if !s.generateInProgress {
		s.closePopup(w)
		return
	}
	if s.handlePopupEscape(w) {
		return
	}

	w.Row(28).Dynamic(1)
	w.Label("Generating hotsheets...", "LC")
	s.renderPopupMessage(w, "Please wait while the reports are processed. Closing this popup will not cancel the generation.", 42)
	w.Row(46).Dynamic(1)
	w.Label("", "LC")
	w.Row(32).Static(0, 110, 0)
	w.Label("", "LC")
	if w.ButtonText("Cancel") {
		s.closePopup(w)
	}
	w.Label("", "LC")
}

// openUpdateAvailablePopup shows the optional update prompt.
func (s *AppState) openUpdateAvailablePopup() {
	s.currentPopup = popupUpdateAvailable
	s.mw.PopupOpen("Update Available", nucular.WindowMovable|nucular.WindowTitle|nucular.WindowDynamic|nucular.WindowNoScrollbar, s.centeredPopupRect(640, 240), true, s.renderUpdateAvailablePopup)
}

// renderUpdateAvailablePopup draws the popup that offers either updating now or
// continuing to use the current version.
func (s *AppState) renderUpdateAvailablePopup(w *nucular.Window) {
	if s.updateInProgress {
		s.closePopup(w)
		return
	}
	if s.handleUpdateAvailablePopupKeyboard(w) {
		return
	}

	message := fmt.Sprintf("A new version (%s) is available. You can update now, or continue using the current version.", s.latestVersion)
	s.renderPopupMessage(w, message, 50)
	w.Row(50).Dynamic(1)
	w.Label("", "LC")
	w.Row(18).Dynamic(1)
	w.Label(fmt.Sprintf("%s to update, %s to continue, Esc to close.", shortcutDisplay("U"), shortcutDisplay("C")), "LC")
	w.Row(32).Static(0, 160, 24, 180, 0)
	w.Label("", "LC")
	if w.ButtonText(buttonShortcutLabel("Update", "U")) {
		s.beginSelfUpdate()
		return
	}
	w.Label("", "LC")
	if w.ButtonText(buttonShortcutLabel("Continue", "C")) {
		s.closePopup(w)
	}
	w.Label("", "LC")
}

// openUpdateProgressPopup shows the modal progress popup while a self-update is
// being applied.
func (s *AppState) openUpdateProgressPopup() {
	s.currentPopup = popupUpdateProgress
	s.mw.PopupOpen("Updating", nucular.WindowMovable|nucular.WindowTitle|nucular.WindowDynamic|nucular.WindowNoScrollbar, s.centeredPopupRect(460, 200), true, s.renderUpdateProgressPopup)
}

// renderUpdateProgressPopup draws the content of the self-update progress
// popup.
func (s *AppState) renderUpdateProgressPopup(w *nucular.Window) {
	if !s.updateInProgress {
		s.closePopup(w)
		return
	}
	if s.handlePopupEscape(w) {
		return
	}

	w.Row(28).Dynamic(1)
	w.Label("Updating application...", "LC")
	s.renderPopupMessage(w, "Please wait while the new version is downloaded and applied. Closing this popup will not stop the update.", 42)
	w.Row(46).Dynamic(1)
	w.Label("", "LC")
	w.Row(32).Static(0, 110, 0)
	w.Label("", "LC")
	if w.ButtonText("Cancel") {
		s.closePopup(w)
	}
	w.Label("", "LC")
}

// openOutputsPopup shows the modal popup that lists the generated output files.
func (s *AppState) openOutputsPopup() {
	s.currentPopup = popupOutputs
	s.mw.PopupOpen("Created Hotsheets", nucular.WindowMovable|nucular.WindowTitle|nucular.WindowDynamic|nucular.WindowNoScrollbar, s.centeredPopupRect(660, 430), true, s.renderOutputsPopup)
}

// renderOutputsPopup draws the output list and action buttons after generation
// completes.
func (s *AppState) renderOutputsPopup(w *nucular.Window) {
	s.handleOutputsPopupKeyboard(w)

	if len(s.outputs) == 0 {
		w.Row(20).Dynamic(1)
		w.Label("No files were created.", "LC")
		s.renderPopupMessage(w, "No output files were returned by the generator.", 52)
		w.Row(172).Dynamic(1)
		w.Label("", "LC")
	} else {
		w.Row(20).Dynamic(1)
		w.Label(fmt.Sprintf("Created files (%d):", len(s.outputs)), "LC")
		w.Row(18).Dynamic(1)
		w.Label(fmt.Sprintf("Double-click to open, use Up/Down and Enter, %s for folder, %s for done, or Esc to close.", shortcutDisplay("O"), shortcutDisplay("D")), "LC")
		w.Row(235).Dynamic(1)
		if gl, gw := nucular.GroupListStart(w, len(s.outputs), "created-hotsheets", nucular.WindowBorder|nucular.WindowNoHScrollbar); gw != nil {
			// SkipToVisible keeps large result lists from rendering every row on
			// every frame while still preserving the current scroll position.
			gl.SkipToVisible(22)
			gw.Row(22).Dynamic(1)
			for gl.Next() {
				idx := gl.Index()
				selected := idx == s.selectedOutput
				if gw.SelectableLabel(s.outputs[idx], "LC", &selected) {
					s.handleOutputClick(idx)
				}
				if idx == s.selectedOutput && s.selectedOutputNeedsScroll {
					gl.Center()
					s.selectedOutputNeedsScroll = false
				}
			}
		}
	}

	w.Row(32).Static(0, 180, 24, 150, 0)
	w.Label("", "LC")
	if w.ButtonText(buttonShortcutLabel("Open Folder", "O")) {
		s.openSelectedOutputFolder()
	}
	w.Label("", "LC")
	if w.ButtonText(buttonShortcutLabel("Done", "D")) {
		s.closeOutputsPopup(w)
	}
	w.Label("", "LC")
}

// handleOutputsPopupKeyboard applies keyboard navigation and activation for the
// generated output list while the popup is open.
func (s *AppState) handleOutputsPopupKeyboard(w *nucular.Window) {
	if s.handlePopupEscape(w) || len(s.outputs) == 0 {
		return
	}

	in := w.Input()
	if in == nil {
		return
	}

	switch {
	case in.Keyboard.Pressed(key.CodeDownArrow):
		s.moveSelectedOutput(1)
	case in.Keyboard.Pressed(key.CodeUpArrow):
		s.moveSelectedOutput(-1)
	case in.Keyboard.Pressed(key.CodeReturnEnter), in.Keyboard.Pressed(key.CodeKeypadEnter):
		s.openSelectedOutput()
	case hasShortcut(in.Keyboard.Keys, key.CodeO):
		s.openSelectedOutputFolder()
	case hasShortcut(in.Keyboard.Keys, key.CodeD):
		s.closeOutputsPopup(w)
	}
}

// closeOutputsPopup dismisses the output-results popup and resets the inputs for
// the next generation run.
func (s *AppState) closeOutputsPopup(w *nucular.Window) {
	s.resetInputs()
	s.outputs = nil
	s.closePopup(w)
}

// handleUpdateAvailablePopupKeyboard applies the update prompt's keyboard
// shortcuts and returns true when one of them closes or advances the popup.
func (s *AppState) handleUpdateAvailablePopupKeyboard(w *nucular.Window) bool {
	if s.handlePopupEscape(w) {
		return true
	}

	in := w.Input()
	if in == nil {
		return false
	}

	switch {
	case hasShortcut(in.Keyboard.Keys, key.CodeU):
		s.beginSelfUpdate()
		return true
	case hasShortcut(in.Keyboard.Keys, key.CodeC):
		s.closePopup(w)
		return true
	}

	return false
}

// handlePopupEscape closes the active popup when Escape is pressed.
func (s *AppState) handlePopupEscape(w *nucular.Window) bool {
	in := w.Input()
	if in == nil {
		return false
	}
	if !in.Keyboard.Pressed(key.CodeEscape) {
		return false
	}
	s.closePopup(w)
	return true
}

// closePopup dismisses the currently open popup and clears the popup state.
func (s *AppState) closePopup(w *nucular.Window) {
	s.currentPopup = popupNone
	w.Close()
}

// centeredPopupRect returns a popup rectangle that is centered inside the main
// window.
//
// Nucular applies style scaling to popup rectangles when PopupOpen is called
// with scale=true. The main window bounds stored in AppState are already scaled,
// so this helper converts them back to unscaled coordinates before centering the
// popup. Without that adjustment the popup drifts toward the lower-right on
// high-DPI layouts and can end up partially off-screen.
func (s *AppState) centeredPopupRect(width, height int) rect.Rect {
	bounds := s.windowBounds
	if bounds.W <= 0 || bounds.H <= 0 {
		return rect.Rect{X: 40, Y: 40, W: width, H: height}
	}

	scale := 1.0
	if s.mw != nil && s.mw.Style() != nil && s.mw.Style().Scaling > 0 {
		scale = s.mw.Style().Scaling
	}

	unscaledX := int(float64(bounds.X) / scale)
	unscaledY := int(float64(bounds.Y) / scale)
	unscaledW := int(float64(bounds.W) / scale)
	unscaledH := int(float64(bounds.H) / scale)
	if unscaledW <= 0 || unscaledH <= 0 {
		return rect.Rect{X: 40, Y: 40, W: width, H: height}
	}

	margin := 20
	maxWidth := unscaledW - 2*margin
	maxHeight := unscaledH - 2*margin
	if maxWidth > 0 && width > maxWidth {
		width = maxWidth
	}
	if maxHeight > 0 && height > maxHeight {
		height = maxHeight
	}

	x := unscaledX + (unscaledW-width)/2
	y := unscaledY + (unscaledH-height)/2
	if x < unscaledX+margin {
		x = unscaledX + margin
	}
	if y < unscaledY+margin {
		y = unscaledY + margin
	}
	return rect.Rect{X: x, Y: y, W: width, H: height}
}

// renderPopupMessage draws popup text line-by-line using the custom word-wrap
// helper.
func (s *AppState) renderPopupMessage(w *nucular.Window, message string, maxChars int) {
	for _, line := range wrapPopupText(message, maxChars) {
		w.Row(24).Dynamic(1)
		w.Label(line, "LC")
	}
}

// wrapPopupText wraps popup copy on word boundaries so the renderer does not
// split words in the middle on narrow lines.
func wrapPopupText(message string, maxChars int) []string {
	if maxChars <= 0 {
		return []string{message}
	}

	paragraphs := strings.Split(message, "\n")
	lines := make([]string, 0, len(paragraphs))
	for idx, paragraph := range paragraphs {
		words := strings.Fields(paragraph)
		if len(words) == 0 {
			lines = append(lines, "")
			continue
		}

		current := words[0]
		for _, word := range words[1:] {
			candidate := current + " " + word
			if utf8.RuneCountInString(candidate) <= maxChars {
				current = candidate
				continue
			}
			lines = append(lines, current)
			current = word
		}
		lines = append(lines, current)

		if idx < len(paragraphs)-1 {
			lines = append(lines, "")
		}
	}
	return lines
}
