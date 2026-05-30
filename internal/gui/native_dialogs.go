package gui

import (
	"errors"
	"fmt"
	"os/exec"
	"runtime"
	"strings"
)

// errDialogCancelled is returned when the user dismisses a native picker without
// selecting a file or directory.
var errDialogCancelled = errors.New("cancelled")

// pickFile opens a platform-native file picker and returns the chosen path.
//
// The implementation intentionally avoids CGO-backed dialog libraries so the
// project can continue to build with `CGO_ENABLED=0`.
func pickFile() (string, error) {
	switch runtime.GOOS {
	case "darwin":
		return runDialogCommand("osascript",
			"-e", `POSIX path of (choose file with prompt "Select a report file")`,
		)
	case "windows":
		return runDialogCommand("powershell", "-NoProfile", "-Command", `[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; $dlg = New-Object System.Windows.Forms.OpenFileDialog; $dlg.Filter = 'Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*'; if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $dlg.FileName }`)
	default:
		// On Linux prefer zenity when available, then fall back to kdialog.
		if path, err := runDialogCommand("zenity", "--file-selection", "--title=Select a report file"); err == nil || errors.Is(err, errDialogCancelled) {
			return path, err
		}
		return runDialogCommand("kdialog", "--getopenfilename", ".")
	}
}

// pickDirectory opens a platform-native directory picker and returns the chosen
// folder path.
func pickDirectory() (string, error) {
	switch runtime.GOOS {
	case "darwin":
		return runDialogCommand("osascript",
			"-e", `POSIX path of (choose folder with prompt "Select an output directory")`,
		)
	case "windows":
		return runDialogCommand("powershell", "-NoProfile", "-Command", `[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; $dlg = New-Object System.Windows.Forms.FolderBrowserDialog; if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $dlg.SelectedPath }`)
	default:
		if path, err := runDialogCommand("zenity", "--file-selection", "--directory", "--title=Select an output directory"); err == nil || errors.Is(err, errDialogCancelled) {
			return path, err
		}
		return runDialogCommand("kdialog", "--getexistingdirectory", ".")
	}
}

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
