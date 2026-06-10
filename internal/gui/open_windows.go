//go:build windows

package gui

import (
	"fmt"
	"path/filepath"
	"syscall"
	"unsafe"
)

// Windows shell constants used by the path-opening helpers.
const shellExecuteSuccessThreshold = 32

// Lazily-resolved Windows shell procedures used by this file.
var procShellExecuteW = shell32DLL.NewProc("ShellExecuteW")

// Windows shell entrypoints.

// OpenPath opens a file or folder using the operating system's default handler.
//
// On Windows this calls `ShellExecuteW` directly so Explorer or the registered
// application handles the path without first spawning `cmd.exe`, which avoids
// flashing a terminal window in GUI workflows.
//
// Any launch error is ignored because there is no actionable recovery path in
// the GUI; the user can try again or open the file manually.
func OpenPath(path string) {
	if err := openPath(path); err != nil {
		return
	}
}

// OpenFolderForFile opens the directory containing the provided file path.
//
// This intentionally opens the parent directory rather than selecting the file
// itself because the rest of the GUI already treats the action as a generic
// "open folder" workflow.
func OpenFolderForFile(path string) {
	if path == "" {
		return
	}
	OpenPath(filepath.Dir(path))
}

// Low-level Windows shell helpers.

// openPath resolves a path through the Windows shell using the default `open`
// verb.
//
// Returning an error from this helper keeps the low-level shell interaction
// isolated and testable while allowing `OpenPath` to preserve the package's
// existing fire-and-forget behavior.
func openPath(path string) error {
	if path == "" {
		return nil
	}

	operation, err := syscall.UTF16PtrFromString("open")
	if err != nil {
		return fmt.Errorf("failed to encode shell operation: %w", err)
	}
	target, err := syscall.UTF16PtrFromString(path)
	if err != nil {
		return fmt.Errorf("failed to encode shell target: %w", err)
	}

	// Pass a nil HWND and nil parameter/directory pointers because this helper is
	// only asking the shell to perform its default "open" action for the target
	// path. The final argument uses `SW_SHOWNORMAL` semantics so Explorer or the
	// registered application decides how to present the window.
	result, _, _ := procShellExecuteW.Call(0, uintptr(unsafe.Pointer(operation)), uintptr(unsafe.Pointer(target)), 0, 0, 1)
	if result <= shellExecuteSuccessThreshold {
		// Per the ShellExecuteW contract, return values greater than 32 indicate
		// success. Smaller values are shell-defined error codes.
		return fmt.Errorf("ShellExecuteW failed with code %d", result)
	}
	return nil
}
