package gui

import (
	"fmt"
	"strings"

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
	case hasShortcut(in.Keyboard.Keys, key.CodeD) && !s.isBusy():
		s.browseOutputDir()
	case hasShortcut(in.Keyboard.Keys, key.CodeQ):
		s.quit()
	case hasShortcut(in.Keyboard.Keys, key.CodeU) && !s.isBusy() && !s.updateCheckInProgress:
		s.startUpdateCheck(true)
	case hasShortcut(in.Keyboard.Keys, key.CodeG) && !s.isBusy() && !s.updateCheckInProgress:
		s.startGenerate()
	}
}

// buttonShortcutLabel formats a button title with the platform-appropriate
// shortcut glyph/text.
func buttonShortcutLabel(labelText, keyName string) string {
	return fmt.Sprintf("%s (%s)", labelText, shortcutDisplay(keyName))
}

// shortcutHint appends a shortcut hint to a muted line of helper text.
func shortcutHint(baseText, keyName string) string {
	return fmt.Sprintf("%s • %s", baseText, shortcutDisplay(keyName))
}

// shortcutDisplay renders the application shortcut modifier in a user-facing
// form appropriate for the current platform.
func shortcutDisplay(keyName string) string {
	keyName = strings.ToUpper(strings.TrimSpace(keyName))
	if shortcutModifier() == key.ModMeta {
		return "Cmd+" + keyName
	}
	return "Ctrl+" + keyName
}

// hasShortcut reports whether the current keyboard batch contains the given app
// shortcut using the platform shortcut modifier and no extra modifiers beyond
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
