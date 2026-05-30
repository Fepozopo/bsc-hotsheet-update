package gui

import (
	"os/exec"
	"path/filepath"
	"runtime"
)

// OpenPath opens a file or folder using the platform's default handler.
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
