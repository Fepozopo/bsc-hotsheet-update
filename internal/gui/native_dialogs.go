package gui

import (
	"errors"
	"fmt"
	"os/exec"
	"runtime"
	"strings"
)

var errDialogCancelled = errors.New("cancelled")

func pickFile() (string, error) {
	switch runtime.GOOS {
	case "darwin":
		return runDialogCommand("osascript",
			"-e", `POSIX path of (choose file with prompt "Select a report file")`,
		)
	case "windows":
		return runDialogCommand("powershell", "-NoProfile", "-Command", `[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; $dlg = New-Object System.Windows.Forms.OpenFileDialog; $dlg.Filter = 'Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*'; if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $dlg.FileName }`)
	default:
		if path, err := runDialogCommand("zenity", "--file-selection", "--title=Select a report file"); err == nil || errors.Is(err, errDialogCancelled) {
			return path, err
		}
		return runDialogCommand("kdialog", "--getopenfilename", ".")
	}
}

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

func isDialogCancelled(message string) bool {
	msg := strings.ToLower(strings.TrimSpace(message))
	if msg == "" {
		return false
	}
	return strings.Contains(msg, "cancel") || strings.Contains(msg, "canceled") || strings.Contains(msg, "cancelled")
}
