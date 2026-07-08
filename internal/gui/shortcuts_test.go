package gui

import (
	"testing"

	"golang.org/x/mobile/event/key"
)

// TestShortcutModifierUsesAlt verifies that every application shortcut now uses
// the shared Option/Alt modifier instead of platform-specific Command/Control
// bindings.
func TestShortcutModifierUsesAlt(t *testing.T) {
	if got := shortcutModifier(); got != key.ModAlt {
		t.Fatalf("shortcutModifier() = %v, want %v", got, key.ModAlt)
	}
}

// TestShortcutLabelMarksMnemonic verifies that visible shortcut cues are
// rendered by wrapping the matching mnemonic character in brackets instead of
// appending a modifier string such as Cmd+I or Ctrl+I.
func TestShortcutLabelMarksMnemonic(t *testing.T) {
	got := shortcutLabel("Output Directory (optional):", "O")
	want := "[O]utput Directory (optional):"
	if got != want {
		t.Fatalf("shortcutLabel() = %q, want %q", got, want)
	}
}

// TestHasShortcutAllowsShift verifies that uppercase key presses still trigger
// a shortcut because Shift is commonly used to produce capital letters while
// holding the shared Option/Alt modifier.
func TestHasShortcutAllowsShift(t *testing.T) {
	keys := []key.Event{{Code: key.CodeP, Modifiers: key.ModAlt | key.ModShift}}
	if !hasShortcut(keys, key.CodeP) {
		t.Fatal("hasShortcut() = false, want true for Alt+Shift+P")
	}
}

// TestHasShortcutRejectsExtraModifiers verifies that only the shared Option/Alt
// modifier, plus an optional Shift key, can activate an application shortcut.
func TestHasShortcutRejectsExtraModifiers(t *testing.T) {
	keys := []key.Event{{Code: key.CodeP, Modifiers: key.ModAlt | key.ModControl}}
	if hasShortcut(keys, key.CodeP) {
		t.Fatal("hasShortcut() = true, want false when an extra modifier is present")
	}
}
