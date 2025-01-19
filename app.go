package main

import (
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
)

func openFileWindow(parent fyne.Window, callback func(r fyne.URIReadCloser, e error)) {
	dialog.NewFileOpen(func(r fyne.URIReadCloser, e error) {
		callback(r, e)
	}, parent).Show()
}

func selectFiles(a fyne.App) (string, string, string, string) {
	window := a.NewWindow("Hotsheet Updater")
	window.SetContent(widget.NewLabel("Please select the files to update:"))
	window.Resize(fyne.NewSize(900, 800))

	files := make([]*widget.Entry, 3)
	buttons := make([]*widget.Button, 3)

	options := []string{"smd", "bsc", "21c"}
	list := widget.NewSelect(options, func(s string) {
	})

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
	var filePaths []string
	submitButton := widget.NewButton("Submit", func() {
		selection = list.Selected
		filePaths = make([]string, len(files))
		for i, entry := range files {
			filePaths[i] = entry.Text
		}
		window.Close()
	})

	window.SetContent(container.New(
		layout.NewVBoxLayout(),
		widget.NewLabelWithStyle("Which hotsheet would you like to update?", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		list,
		layout.NewSpacer(),
		widget.NewLabelWithStyle("Select Hotsheet:", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		files[0],
		buttons[0],
		layout.NewSpacer(),
		widget.NewLabelWithStyle("Select Stock Report:", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		files[1],
		buttons[1],
		layout.NewSpacer(),
		widget.NewLabelWithStyle("Select Sales Report:", fyne.TextAlignCenter, fyne.TextStyle{Bold: true}),
		files[2],
		buttons[2],
		layout.NewSpacer(),
		submitButton,
	))

	window.ShowAndRun()

	return selection, filePaths[0], filePaths[1], filePaths[2]
}
