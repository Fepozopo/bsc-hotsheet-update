package gui

import (
	"os/exec"
	"path/filepath"
	"runtime"
)

// OpenPath opens a file or folder using the operating system's default handler.
//
// Any launch error is ignored because there is no actionable recovery path in
// the GUI; the user can try again or open the file manually.
func OpenPath(path string) {
	if path == "" {
		return
	}

	switch runtime.GOOS {
	case "darwin":
		_ = exec.Command("open", path).Start()
	case "windows":
		_ = exec.Command("cmd", "/C", "start", "", path).Start()
	default:
		_ = exec.Command("xdg-open", path).Start()
	}
}

// OpenFolderForFile opens the directory containing the provided file path.
func OpenFolderForFile(path string) {
	if path == "" {
		return
	}
	OpenPath(filepath.Dir(path))
}
