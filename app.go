package main

import (
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
)

// openFileWindow creates a file open dialog and calls the given callback function with the selected file.
// If the user cancels the dialog, the error argument will be set to an error with message "cancelled".
func openFileWindow(parent fyne.Window, callback func(r fyne.URIReadCloser, e error)) {
	dialog.NewFileOpen(func(r fyne.URIReadCloser, e error) {
		callback(r, e)
	}, parent).Show()
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
				openFileWindow(window, func(r fyne.URIReadCloser, e error) {
					if e != nil {
						return
					}
					files[i].SetText(r.URI().Path())
				})
			}
		}(i))
	}

	var selection string
	var hotsheetPaths []string

	hotsheetLabels := []string{"21c Hotsheet:", "BSC Hotsheet:", "BJP Hotsheet:", "SMD Hotsheet:"}

	hotsheetLabel := widget.NewLabelWithStyle("Select Hotsheet:", fyne.TextAlignCenter, fyne.TextStyle{Bold: true})
	hotsheetRows := []fyne.CanvasObject{
		hotsheetLabel,
		files[0],
		buttons[0],
	}
	// Add all four hotsheet fields (21C, BSC, BJP, SMD)
	for i := 0; i < 4; i++ {
		hotsheetRows = append(hotsheetRows, widget.NewLabelWithStyle(hotsheetLabels[i], fyne.TextAlignCenter, fyne.TextStyle{Bold: true}))
		hotsheetRows = append(hotsheetRows, files[i], buttons[i])
	}

	// Start with only the label and select list visible
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
			hotsheetPaths = []string{files[0].Text}
		}
		window.Close()
	})

	// Dynamically add fields after selection
	list.OnChanged = func(s string) {
		// Remove all but the label and select list
		content.Objects = content.Objects[:2]
		if s == "All" {
			hotsheetLabel.SetText("Select Hotsheets:")
			// Add all hotsheet fields (21C, BSC, BJP, SMD)
			for i := 0; i < 4; i++ {
				content.Add(hotsheetRows[3+i*3])   // label
				content.Add(hotsheetRows[3+i*3+1]) // entry
				content.Add(hotsheetRows[3+i*3+2]) // button
			}
		} else if s != "" {
			hotsheetLabel.SetText("Select Hotsheet:")
			// Only show the first hotsheet field
			content.Add(hotsheetRows[0])
			content.Add(hotsheetRows[1])
			content.Add(hotsheetRows[2])
		}
		// Always add report fields and submit button if a selection is made
		if s != "" {
			content.Add(layout.NewSpacer())
			content.Add(widget.NewLabelWithStyle("Select Inventory Report:", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}))
			content.Add(files[4])
			content.Add(buttons[4])
			content.Add(layout.NewSpacer())
			content.Add(widget.NewLabelWithStyle("Select PO Report:", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}))
			content.Add(files[5])
			content.Add(buttons[5])
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

	if selection == "All" {
		return selection, hotsheetPaths, files[4].Text, files[5].Text
	} else {
		return selection, hotsheetPaths, files[4].Text, files[5].Text
	}
}
