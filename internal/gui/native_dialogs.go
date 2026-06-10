package gui

import (
	"errors"
	"fmt"
	"os/exec"
	"strings"
)

// errDialogCancelled is returned when the user dismisses a native picker without
// selecting a file or directory.
var errDialogCancelled = errors.New("cancelled")

// runDialogCommand executes a platform-specific helper command and converts its
// output into either a selected path or a user-cancelled result.
func runDialogCommand(name string, args ...string) (string, error) {
	cmd := exec.Command(name, args...)
	out, err := cmd.CombinedOutput()
	text := strings.TrimSpace(string(out))
	if err != nil {
		if isDialogCancelled(text) {
			return "", errDialogCancelled
		}
		if text == "" {
			return "", fmt.Errorf("failed to open native dialog with %s: %w", name, err)
		}
		return "", fmt.Errorf("%s: %s", name, text)
	}
	if text == "" {
		return "", errDialogCancelled
	}
	return text, nil
}

// isDialogCancelled reports whether the dialog helper's output appears to mean
// that the user dismissed the picker without choosing a result.
func isDialogCancelled(message string) bool {
	msg := strings.ToLower(strings.TrimSpace(message))
	if msg == "" {
		return false
	}
	return strings.Contains(msg, "cancel") || strings.Contains(msg, "canceled") || strings.Contains(msg, "cancelled")
}
