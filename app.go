package main

import (
	"errors"
	"fmt"
	"os"
	"os/exec"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
	"github.com/Fepozopo/bsc-hotsheet-update/hotsheet"
	"github.com/Fepozopo/bsc-hotsheet-update/internal/version"
	"github.com/blang/semver"
	"github.com/rhysd/go-github-selfupdate/selfupdate"

	osDialog "github.com/sqweek/dialog"
)

// openFileWindow creates a file open dialog using the system's native file manager
// and calls the given callback function with the selected file path.
// If the user cancels the dialog, the error argument will be set to an error with message "cancelled".
func openFileWindow(parent fyne.Window, callback func(filePath string, e error)) {
	filePath, err := osDialog.File().Load() // Use the aliased dialog for the native file open
	if err != nil {
		if err.Error() == "cancelled" {
			callback("", errors.New("cancelled"))
		} else {
			callback("", err)
		}
		return
	}
	callback(filePath, nil)
}

// openDirWindow opens a native directory chooser and returns the chosen directory path
func openDirWindow(parent fyne.Window, callback func(dirPath string, e error)) {
	dirPath, err := osDialog.Directory().Browse()
	if err != nil {
		if err.Error() == "cancelled" {
			callback("", errors.New("cancelled"))
		} else {
			callback("", err)
		}
		return
	}
	callback(dirPath, nil)
}

func checkForUpdates(w fyne.Window, showNoUpdatesDialog bool) {
	go func() {
		const repo = "Fepozopo/bsc-hotsheet-update"
		latest, found, err := selfupdate.DetectLatest(repo)
		if err != nil {
			dialog.ShowError(fmt.Errorf("update check failed: %w", err), w)
			return
		}

		currentVer, _ := semver.Parse(version.Version)
		if !found || latest.Version.Equals(currentVer) {
			if showNoUpdatesDialog {
				dialog.ShowInformation("No Updates", "You are already running the latest version.", w)
			}
			return
		}
		updateMsg := fmt.Sprintf("A new version (%s) is available. You must update to continue using the application.", latest.Version)
		dialog.NewCustomConfirm(
			"Update Required",
			"Update",
			"Quit",
			widget.NewLabel(updateMsg),
			func(ok bool) {
				if ok {
					exe, err := os.Executable()
					if err != nil {
						dialog.ShowError(fmt.Errorf("could not locate executable: %w", err), w)
						return
					}

					// Show infinite progress bar dialog
					progress := widget.NewProgressBarInfinite()
					progressLabel := widget.NewLabel("Updating application...")
					progressDialog := dialog.NewCustom("Updating", "Cancel", container.NewVBox(progressLabel, progress), w)
					progressDialog.Show()

					go func() {
						err = selfupdate.UpdateTo(latest.AssetURL, exe)
						progressDialog.Hide()
						if err != nil {
							dialog.ShowError(fmt.Errorf("update failed: %w", err), w)
							return
						}
						// Force restart
						cmd := exec.Command(exe, os.Args[1:]...)
						cmd.Env = os.Environ()
						err := cmd.Start()
						if err != nil {
							dialog.ShowError(fmt.Errorf("failed to restart: %w", err), w)
							return
						}
						os.Exit(0)
					}()
				} else {
					os.Exit(0)
				}
			},
			w,
		).Show()
	}()
}

// selectFiles creates a GUI window that asks for the required reports and output directory,
// but does the hotsheet generation itself. When generation finishes it opens a Fyne window
// listing the created files. The main window is not closed after generation; instead the
// inputs are cleared so the user can run another generation.
func selectFiles(a fyne.App) (string, string, string) {
	window := a.NewWindow("Hotsheet Generator")
	checkForUpdates(window, false)
	window.Resize(fyne.NewSize(700, 420))

	// Entries for reports and output
	inventoryEntry := widget.NewEntry()
	inventoryEntry.SetPlaceHolder("Path to inventory report (xlsx)")
	poEntry := widget.NewEntry()
	poEntry.SetPlaceHolder("Path to PO report (xlsx) (optional)")
	outputEntry := widget.NewEntry()
	outputEntry.SetPlaceHolder("Output directory (optional)")

	// Browse buttons
	invBtn := widget.NewButton("Browse", func() {
		openFileWindow(window, func(filePath string, e error) {
			if e != nil {
				if e.Error() == "cancelled" {
					return
				}
				dialog.ShowError(e, window)
				return
			}
			inventoryEntry.SetText(filePath)
		})
	})
	poBtn := widget.NewButton("Browse", func() {
		openFileWindow(window, func(filePath string, e error) {
			if e != nil {
				if e.Error() == "cancelled" {
					return
				}
				dialog.ShowError(e, window)
				return
			}
			poEntry.SetText(filePath)
		})
	})
	outBtn := widget.NewButton("Browse", func() {
		openDirWindow(window, func(dirPath string, e error) {
			if e != nil {
				if e.Error() == "cancelled" {
					return
				}
				dialog.ShowError(e, window)
				return
			}
			outputEntry.SetText(dirPath)
		})
	})

	// Generate handler - performs generation in background and shows outputs in a new window
	generate := func() {
		if strings.TrimSpace(inventoryEntry.Text) == "" {
			dialog.ShowError(errors.New("Inventory report is required"), window)
			return
		}

		// Progress dialog while generating
		progress := widget.NewProgressBarInfinite()
		progressLabel := widget.NewLabel("Generating hotsheets...")
		progressDialog := dialog.NewCustom("Generating Hotsheets", "Cancel", container.NewVBox(progressLabel, progress), window)
		progressDialog.Show()

		// Run generation in goroutine to avoid blocking UI
		go func(inv, po, outdir string) {
			outputs, err := hotsheet.CreateFromReports(inv, po, outdir)
			// Must manipulate UI from main goroutine; schedule with fyne.CurrentApp().SendNotification isn't appropriate here,
			// but dialog.Show* and window operations are safe to call from other goroutines in Fyne as they marshal to the main loop.
			progressDialog.Hide()
			if err != nil {
				// Show error on main window
				dialog.ShowError(err, window)
				return
			}

			// Prepare outputs window
			outWin := a.NewWindow("Created Hotsheets")
			outWin.Resize(fyne.NewSize(600, 400))

			list := widget.NewList(
				func() int { return len(outputs) },
				func() fyne.CanvasObject { return widget.NewLabel("") },
				func(i widget.ListItemID, o fyne.CanvasObject) {
					o.(*widget.Label).SetText(outputs[i])
				},
			)

			// If there are no outputs, show a message
			var content fyne.CanvasObject
			if len(outputs) == 0 {
				content = container.NewVBox(widget.NewLabel("No files were created."))
			} else {
				// Put the label in the top border and make the list scrollable so it expands to fill available space.
				// Using NewVScroll(list) allows the list to take the remaining height while the window expands.
				content = container.NewBorder(widget.NewLabel("Created files:"), nil, nil, nil, container.NewVScroll(list))
			}

			doneBtn := widget.NewButton("Done", func() {
				outWin.Close()
				// Clear the main window fields so nothing is selected
				inventoryEntry.SetText("")
				poEntry.SetText("")
				outputEntry.SetText("")
			})

			// Use a border so the Done button stays at the bottom and the content fills the middle area.
			outWin.SetContent(container.NewBorder(nil, doneBtn, nil, nil, content))
			outWin.Show()
		}(inventoryEntry.Text, poEntry.Text, outputEntry.Text)
	}

	// Buttons: Generate and Quit (Quit closes the main window)
	submitBtn := widget.NewButton("Generate Hotsheets", generate)
	quitBtn := widget.NewButton("Quit", func() {
		window.Close()
	})

	buttons := container.NewHBox(layout.NewSpacer(), submitBtn, widget.NewLabel("   "), quitBtn)

	content := container.NewVBox(
		widget.NewLabelWithStyle("Create Unified Hotsheets from Reports", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		layout.NewSpacer(),
		widget.NewLabelWithStyle("Inventory Report:", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		container.NewBorder(nil, nil, invBtn, nil, inventoryEntry),
		layout.NewSpacer(),
		widget.NewLabelWithStyle("PO Report (optional):", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		container.NewBorder(nil, nil, poBtn, nil, poEntry),
		layout.NewSpacer(),
		widget.NewLabelWithStyle("Output Directory (optional):", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		container.NewBorder(nil, nil, outBtn, nil, outputEntry),
		layout.NewSpacer(),
		buttons,
	)

	window.SetContent(content)
	window.ShowAndRun()

	// Return empty strings because the generation is handled inside this UI flow.
	// This prevents the main() caller from trying to re-run generation.
	return "", "", ""
}
