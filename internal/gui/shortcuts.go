package gui

import (
	"strings"
	"unicode"

	"github.com/aarzilli/nucular"
	"golang.org/x/mobile/event/key"
)

// handleMainKeyboard applies application-level shortcuts when the main window is
// active and no popup currently owns the interaction flow.
func (s *AppState) handleMainKeyboard(w *nucular.Window) {
	if s.currentPopup != popupNone || s.anyEditorActive() {
		return
	}

	in := w.Input()
	if in == nil {
		return
	}

	switch {
	case hasShortcut(in.Keyboard.Keys, key.CodeI) && !s.isBusy():
		s.browseInventory()
	case hasShortcut(in.Keyboard.Keys, key.CodeP) && !s.isBusy():
		s.browsePO()
	case hasShortcut(in.Keyboard.Keys, key.CodeO) && !s.isBusy():
		s.browseOutputDir()
	case hasShortcut(in.Keyboard.Keys, key.CodeQ):
		s.quit()
	case hasShortcut(in.Keyboard.Keys, key.CodeU) && !s.isBusy() && !s.updateCheckInProgress:
		s.startUpdateCheck(true)
	case hasShortcut(in.Keyboard.Keys, key.CodeG) && !s.isBusy() && !s.updateCheckInProgress:
		s.startGenerate()
	}
}

// shortcutLabel wraps the mnemonic character used by a shortcut in brackets so
// the UI can hint at the key without rendering the full modifier combination.
//
// Bracketed mnemonics render consistently across the current text stack,
// including button labels that only support plain text.
func shortcutLabel(labelText, keyName string) string {
	keyName = strings.TrimSpace(keyName)
	if keyName == "" {
		return labelText
	}

	keyRune := []rune(strings.ToLower(keyName))[0]
	var builder strings.Builder
	matched := false
	for _, r := range labelText {
		// Only the first matching rune is marked so each label exposes a single
		// mnemonic cue instead of visually cluttering repeated letters.
		if !matched && unicode.ToLower(r) == keyRune {
			builder.WriteRune('[')
			builder.WriteRune(r)
			builder.WriteRune(']')
			matched = true
			continue
		}
		builder.WriteRune(r)
	}
	return builder.String()
}

// buttonShortcutLabel formats a button title by marking the mnemonic character
// associated with the button shortcut.
func buttonShortcutLabel(labelText, keyName string) string {
	return shortcutLabel(labelText, keyName)
}

// hasShortcut reports whether the current keyboard batch contains the given app
// shortcut using the shared Option/Alt modifier and no extra modifiers beyond
// an optional Shift key.
func hasShortcut(keys []key.Event, code key.Code) bool {
	required := shortcutModifier()
	for _, k := range keys {
		if k.Code != code {
			continue
		}
		if k.Modifiers&required == 0 {
			continue
		}
		if k.Modifiers&^(required|key.ModShift) != 0 {
			continue
		}
		return true
	}
	return false
}
