package main

import (
	"errors"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	fyneDialog "fyne.io/fyne/v2/dialog" // Alias Fyne's dialog package
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"

	"github.com/sqweek/dialog" // Keep this as 'dialog' since it's the one we'll primarily use for file selection
)

// openFileWindow creates a file open dialog using the system's native file manager
// and calls the given callback function with the selected file path.
// If the user cancels the dialog, the error argument will be set to an error with message "cancelled".
func openFileWindow(parent fyne.Window, callback func(filePath string, e error)) {
	filePath, err := dialog.File().Load() // Use the aliased dialog for the native file open
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

// selectFiles creates a GUI window to select the product line to update and the paths to the hotsheet, stock report, and sales report files.
// It then returns the selection and the paths as strings.
func selectFiles(a fyne.App) (string, []string, string, string) {
	window := a.NewWindow("Hotsheet Updater")
	window.SetContent(widget.NewLabel("Please select the files to update:"))
	window.Resize(fyne.NewSize(900, 800))

	files := make([]*widget.Entry, 6) // 4 hotsheets + 2 reports
	buttons := make([]*widget.Button, 6)

	options := []string{"All", "21c", "BSC", "BJP", "SMD"}
	list := widget.NewSelect(options, nil)

	for i := range files {
		files[i] = widget.NewEntry()
		buttons[i] = widget.NewButton("Browse", func(i int) func() {
			return func() {
				openFileWindow(window, func(filePath string, e error) {
					if e != nil {
						if e.Error() == "cancelled" {
							// User cancelled, no action needed other than not setting the path
						} else {
							fyneDialog.ShowError(e, window) // Use the aliased Fyne dialog for error messages
						}
						return
					}
					files[i].SetText(filePath)
				})
			}
		}(i))
	}

	var selection string
	var hotsheetPaths []string
	var inventoryReportPath string
	var poReportPath string

	hotsheetLabels := []string{"21c Hotsheet:", "BSC Hotsheet:", "BJP Hotsheet:", "SMD Hotsheet:"}

	hotsheetLabel := widget.NewLabelWithStyle("Select Hotsheet:", fyne.TextAlignCenter, fyne.TextStyle{Bold: true})
	hotsheetRows := []fyne.CanvasObject{
		hotsheetLabel,
	}
	for i := 0; i < 4; i++ {
		hotsheetRows = append(hotsheetRows, widget.NewLabelWithStyle(hotsheetLabels[i], fyne.TextAlignCenter, fyne.TextStyle{Bold: true}))
		hotsheetRows = append(hotsheetRows, files[i], buttons[i])
	}
	hotsheetRows = append(hotsheetRows, widget.NewLabelWithStyle("Select Inventory Report:", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}))
	hotsheetRows = append(hotsheetRows, files[4], buttons[4])
	hotsheetRows = append(hotsheetRows, widget.NewLabelWithStyle("Select PO Report:", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}))
	hotsheetRows = append(hotsheetRows, files[5], buttons[5])

	content := container.NewVBox(
		widget.NewLabelWithStyle("Which hotsheet would you like to update? (Select 'All' to update all 4)", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		list,
	)

	submitButton := widget.NewButton("Submit", func() {
		selection = list.Selected
		if selection == "All" {
			hotsheetPaths = make([]string, 4)
			for i := 0; i < 4; i++ {
				hotsheetPaths[i] = files[i].Text
			}
		} else {
			selectedIndex := -1
			for idx, opt := range options {
				if opt == selection {
					selectedIndex = idx - 1
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
		window.Close()
	})

	list.OnChanged = func(s string) {
		content.Objects = content.Objects[:2]

		if s == "All" {
			for i := 0; i < 4; i++ {
				content.Add(hotsheetRows[1+i*3])
				content.Add(hotsheetRows[1+i*3+1])
				content.Add(hotsheetRows[1+i*3+2])
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
				content.Add(hotsheetRows[1+selectedIndex*3])
				content.Add(hotsheetRows[1+selectedIndex*3+1])
				content.Add(hotsheetRows[1+selectedIndex*3+2])
			}
		}

		if s != "" {
			content.Add(layout.NewSpacer())
			content.Add(hotsheetRows[1+4*3])
			content.Add(hotsheetRows[1+4*3+1])
			content.Add(hotsheetRows[1+4*3+2])
			content.Add(layout.NewSpacer())
			content.Add(hotsheetRows[1+5*3])
			content.Add(hotsheetRows[1+5*3+1])
			content.Add(hotsheetRows[1+5*3+2])
			content.Add(layout.NewSpacer())
			content.Add(submitButton)
		}
		content.Refresh()
	}

	window.SetContent(content)
	window.ShowAndRun()

	window.SetCloseIntercept(func() {
		window.Close()
	})

	return selection, hotsheetPaths, inventoryReportPath, poReportPath
}
