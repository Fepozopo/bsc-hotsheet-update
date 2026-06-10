//go:build !windows

package gui

import (
	"errors"
	"runtime"
)

// pickFile opens a platform-native file picker and returns the chosen path.
//
// The implementation intentionally avoids CGO-backed dialog libraries so the
// project can continue to build with `CGO_ENABLED=0`.
func pickFile() (string, error) {
	switch runtime.GOOS {
	case "darwin":
		return pickFileDarwin()
	default:
		return pickFileLinux()
	}
}

// pickDirectory opens a platform-native directory picker and returns the chosen
// folder path.
func pickDirectory() (string, error) {
	switch runtime.GOOS {
	case "darwin":
		return pickDirectoryDarwin()
	default:
		return pickDirectoryLinux()
	}
}

// pickFileDarwin opens the macOS file chooser through AppleScript.
func pickFileDarwin() (string, error) {
	return runDialogCommand("osascript",
		"-e", `POSIX path of (choose file with prompt "Select a report file")`,
	)
}

// pickDirectoryDarwin opens the macOS folder chooser through AppleScript.
func pickDirectoryDarwin() (string, error) {
	return runDialogCommand("osascript",
		"-e", `POSIX path of (choose folder with prompt "Select an output directory")`,
	)
}

// pickFileLinux opens a Linux file chooser, preferring zenity and falling back
// to kdialog when necessary.
func pickFileLinux() (string, error) {
	if path, err := runDialogCommand("zenity", "--file-selection", "--title=Select a report file"); err == nil || errors.Is(err, errDialogCancelled) {
		return path, err
	}
	return runDialogCommand("kdialog", "--getopenfilename", ".")
}

// pickDirectoryLinux opens a Linux directory chooser, preferring zenity and
// falling back to kdialog when necessary.
func pickDirectoryLinux() (string, error) {
	if path, err := runDialogCommand("zenity", "--file-selection", "--directory", "--title=Select an output directory"); err == nil || errors.Is(err, errDialogCancelled) {
		return path, err
	}
	return runDialogCommand("kdialog", "--getexistingdirectory", ".")
}
