package main

import (
	"errors"
	"fmt"
	"os"
	"os/exec"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
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

// selectFiles creates a GUI window that asks for which hotsheet(s) to update first and then
// asks for report files. The reports section only appears after the user selects a hotsheet file
// and clicks Next. The Next button is enabled only when required hotsheet file(s) are filled.
func selectFiles(a fyne.App) (string, []string, string, string, string) {
	window := a.NewWindow("Hotsheet Updater")
	checkForUpdates(window, false)
	window.Resize(fyne.NewSize(900, 800))

	// 4 hotsheets + 3 reports
	files := make([]*widget.Entry, 7)
	buttons := make([]*widget.Button, 7)

	options := []string{"All", "21c", "BJP", "BSC", "SMD"}
	list := widget.NewSelect(options, nil)

	// create entries and browse buttons (capture index correctly)
	for i := range files {
		files[i] = widget.NewEntry()
		idx := i
		buttons[i] = widget.NewButton("Browse", func() {
			openFileWindow(window, func(filePath string, e error) {
				if e != nil {
					if e.Error() == "cancelled" {
						// user cancelled - do nothing
						return
					}
					dialog.ShowError(e, window)
					return
				}
				files[idx].SetText(filePath)
				// Trigger OnChanged manually if present
				if files[idx].OnChanged != nil {
					files[idx].OnChanged(filePath)
				}
			})
		})
	}

	var selection string
	var hotsheetPaths []string
	var inventoryReportPath string
	var poReportPath string
	var bnReportPath string

	hotsheetLabels := []string{"21c Hotsheet:", "BJP Hotsheet:", "BSC Hotsheet:", "SMD Hotsheet:"}

	// Build hotsheet section objects
	hotsheetSection := []fyne.CanvasObject{
		widget.NewLabelWithStyle("Select Hotsheet(s):", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
	}
	for i := 0; i < 4; i++ {
		hotsheetSection = append(hotsheetSection,
			widget.NewLabelWithStyle(hotsheetLabels[i], fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
			files[i],
			buttons[i],
		)
	}

	// Report labels/controls (built separately and only added after Next)
	reportHeader := widget.NewLabelWithStyle("(Inventory and PO required, BN optional):", fyne.TextAlignCenter, fyne.TextStyle{Bold: true})
	inventoryLabel := widget.NewLabelWithStyle("Inventory Report:", fyne.TextAlignCenter, fyne.TextStyle{Bold: true})
	poLabel := widget.NewLabelWithStyle("PO Report:", fyne.TextAlignCenter, fyne.TextStyle{Bold: true})
	bnLabel := widget.NewLabelWithStyle("BN Report:", fyne.TextAlignCenter, fyne.TextStyle{Bold: true})
	bnCheck := widget.NewCheck("Include BN report (optional)", nil)

	// Submit button (will be shown in reports view)
	submitButton := widget.NewButton("Submit", func() {
		selection = list.Selected
		if selection == "All" {
			// All updates except BJP as per previous behavior (21c index 0, BSC index 2, SMD index 3)
			hotsheetPaths = []string{files[0].Text, files[2].Text, files[3].Text}
		} else {
			selectedIndex := -1
			for idx, opt := range options {
				if opt == selection {
					selectedIndex = idx - 1 // because options has "All" at start
					break
				}
			}
			if selectedIndex >= 0 && selectedIndex < 4 {
				hotsheetPaths = []string{files[selectedIndex].Text}
			} else {
				hotsheetPaths = []string{}
			}
		}

		inventoryReportPath = files[4].Text
		poReportPath = files[5].Text
		if bnCheck.Checked {
			bnReportPath = files[6].Text
		} else {
			bnReportPath = ""
		}

		// Validation
		if inventoryReportPath == "" {
			dialog.ShowError(errors.New("Inventory report is required"), window)
			return
		}
		if poReportPath == "" {
			dialog.ShowError(errors.New("PO report is required"), window)
			return
		}
		if bnCheck.Checked && bnReportPath == "" {
			dialog.ShowError(errors.New("BN report selected to be included but no file chosen"), window)
			return
		}

		window.Close()
	})

	// Build main content container (title + select)
	content := container.NewVBox(
		widget.NewLabelWithStyle("Which hotsheet would you like to update? (Select 'All' to update all (excluding BJP))", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		list,
	)

	// declare nextButton variable so addHotsheetRows can reference it before assignment below
	var nextButton *widget.Button

	// addHotsheetRows is used to append the appropriate hotsheet rows based on selection
	addHotsheetRows := func(s string) {
		// Ensure content only has header + select before adding
		content.Objects = content.Objects[:2]
		if s == "All" {
			toShow := []int{0, 2, 3}
			for _, i := range toShow {
				content.Add(hotsheetSection[1+i*3])
				content.Add(hotsheetSection[1+i*3+1])
				content.Add(hotsheetSection[1+i*3+2])
			}
		} else if s != "" {
			selectedIndex := -1
			for idx, opt := range options {
				if opt == s {
					selectedIndex = idx - 1
					break
				}
			}
			if selectedIndex >= 0 && selectedIndex < 4 {
				content.Add(hotsheetSection[1+selectedIndex*3])
				content.Add(hotsheetSection[1+selectedIndex*3+1])
				content.Add(hotsheetSection[1+selectedIndex*3+2])
			}
		}
		// after showing hotsheet entries, show spacer and Next button (Next enabled only if required files filled)
		if s != "" {
			content.Add(layout.NewSpacer())
			content.Add(nextButton)
		}
	}

	// Next button: initially disabled until required hotsheet file(s) are filled for the selection
	nextButton = widget.NewButton("Next: Reports", func() {
		// When Next is pressed, hide the hotsheet selection rows and show a smaller title + the reports section
		// Clear the existing content and show a compact Reports title to reduce clutter
		content.Objects = content.Objects[:0]

		// Smaller title replacing the original header and select
		reportTitle := widget.NewLabelWithStyle("Select Report Files", fyne.TextAlignCenter, fyne.TextStyle{Bold: true})
		content.Add(reportTitle)
		content.Add(reportHeader)
		content.Add(layout.NewSpacer())

		// Add reports section
		content.Add(inventoryLabel)
		content.Add(files[4])
		content.Add(buttons[4])

		content.Add(layout.NewSpacer())

		content.Add(poLabel)
		content.Add(files[5])
		content.Add(buttons[5])

		content.Add(layout.NewSpacer())

		content.Add(bnCheck)
		// If checkbox already checked (unlikely at this point) show BN inputs
		if bnCheck.Checked {
			content.Add(bnLabel)
			content.Add(files[6])
			content.Add(buttons[6])
		}

		content.Add(layout.NewSpacer())
		content.Add(submitButton)
		content.Refresh()
	})
	nextButton.Disable()

	// Helper to check whether the required hotsheet file(s) are filled for the current selection
	isHotsheetFilled := func(sel string) bool {
		if sel == "" {
			return false
		}
		if sel == "All" {
			// required: files[0], files[2], files[3]
			return files[0].Text != "" && files[2].Text != "" && files[3].Text != ""
		}
		// find selected index mapping to files[0..3]
		selectedIndex := -1
		for idx, opt := range options {
			if opt == sel {
				selectedIndex = idx - 1
				break
			}
		}
		if selectedIndex >= 0 && selectedIndex < 4 {
			return files[selectedIndex].Text != ""
		}
		return false
	}

	// Add OnChanged handlers to hotsheet entries so that changes enable/disable Next appropriately.
	// Add these now (after nextButton exists) so they can enable/disable the button.
	for i := 0; i < 4; i++ {
		idx := i
		orig := files[idx].OnChanged
		files[idx].OnChanged = func(s string) {
			// call any existing handler
			if orig != nil {
				orig(s)
			}
			// enable Next if the selection's required entries are filled
			if isHotsheetFilled(list.Selected) {
				nextButton.Enable()
			} else {
				nextButton.Disable()
			}
		}
	}

	// When the selection changes, show the relevant hotsheet rows and update Next enabled state
	list.OnChanged = func(s string) {
		// Reset content to header + select and add selected hotsheet rows
		addHotsheetRows(s)
		// Update Next enablement depending on current entries
		if isHotsheetFilled(s) {
			nextButton.Enable()
		} else {
			nextButton.Disable()
		}
		content.Refresh()
	}

	// bnCheck toggles showing the BN file inputs when reports area is visible.
	bnCheck.OnChanged = func(checked bool) {
		// If reports haven't been added yet, nothing to do
		hasReports := false
		for _, obj := range content.Objects {
			if obj == reportHeader {
				hasReports = true
				break
			}
		}
		if !hasReports {
			return
		}
		if checked {
			// insert BN inputs after the checkbox
			idx := -1
			for i, obj := range content.Objects {
				if obj == bnCheck {
					idx = i
					break
				}
			}
			if idx != -1 {
				after := make([]fyne.CanvasObject, 0, len(content.Objects)+3)
				after = append(after, content.Objects[:idx+1]...)
				after = append(after, bnLabel, files[6], buttons[6])
				after = append(after, content.Objects[idx+1:]...)
				content.Objects = after
			}
		} else {
			// remove BN inputs if present
			newObjs := make([]fyne.CanvasObject, 0, len(content.Objects))
			for _, obj := range content.Objects {
				if obj == bnLabel || obj == files[6] || obj == buttons[6] {
					continue
				}
				newObjs = append(newObjs, obj)
			}
			content.Objects = newObjs
		}
		content.Refresh()
	}

	// Ensure the window closes cleanly
	window.SetCloseIntercept(func() {
		window.Close()
	})

	window.SetContent(content)
	window.ShowAndRun()

	return selection, hotsheetPaths, inventoryReportPath, poReportPath, bnReportPath
}
