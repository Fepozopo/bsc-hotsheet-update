package main

import (
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
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

type githubRelease struct {
	TagName string               `json:"tag_name"`
	Assets  []githubReleaseAsset `json:"assets"`
}

type githubReleaseAsset struct {
	Name               string `json:"name"`
	BrowserDownloadURL string `json:"browser_download_url"`
}

func currentReleaseAssetName() string {
	base := "hotsheet"
	switch runtime.GOOS {
	case "windows":
		return fmt.Sprintf("%s-%s-%s.exe", base, runtime.GOOS, runtime.GOARCH)
	default:
		return fmt.Sprintf("%s-%s-%s", base, runtime.GOOS, runtime.GOARCH)
	}
}

func detectLatestRelease(repo string) (semver.Version, string, error) {
	url := fmt.Sprintf("https://api.github.com/repos/%s/releases/latest", repo)
	req, err := http.NewRequest(http.MethodGet, url, nil)
	if err != nil {
		return semver.Version{}, "", err
	}
	req.Header.Set("Accept", "application/vnd.github+json")
	req.Header.Set("X-GitHub-Api-Version", "2022-11-28")
	req.Header.Set("User-Agent", "bsc-hotsheet-update")

	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return semver.Version{}, "", err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(io.LimitReader(resp.Body, 4096))
		return semver.Version{}, "", fmt.Errorf("GitHub releases API returned %s: %s", resp.Status, strings.TrimSpace(string(body)))
	}

	var release githubRelease
	if err := json.NewDecoder(resp.Body).Decode(&release); err != nil {
		return semver.Version{}, "", fmt.Errorf("could not decode GitHub release response: %w", err)
	}

	latestVersion, err := semver.Parse(strings.TrimPrefix(release.TagName, "v"))
	if err != nil {
		return semver.Version{}, "", fmt.Errorf("could not parse release tag %q: %w", release.TagName, err)
	}

	expectedAsset := currentReleaseAssetName()
	for _, asset := range release.Assets {
		if asset.Name == expectedAsset {
			if asset.BrowserDownloadURL == "" {
				return semver.Version{}, "", fmt.Errorf("release asset %q is missing a download URL", expectedAsset)
			}
			return latestVersion, asset.BrowserDownloadURL, nil
		}
	}

	return semver.Version{}, "", fmt.Errorf("release %q does not include asset %q", release.TagName, expectedAsset)
}

func checkForUpdates(w fyne.Window, showNoUpdatesDialog bool) {
	go func() {
		const repo = "Fepozopo/bsc-hotsheet-update"
		latestVersion, latestAssetURL, err := detectLatestRelease(repo)
		if err != nil {
			dialog.ShowError(fmt.Errorf("update check failed: %w", err), w)
			return
		}

		currentVer, _ := semver.Parse(version.Version)
		if !latestVersion.GT(currentVer) {
			if showNoUpdatesDialog {
				dialog.ShowInformation("No Updates", "You are already running the latest version.", w)
			}
			return
		}
		updateMsg := fmt.Sprintf("A new version (%s) is available. You must update to continue using the application.", latestVersion)
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
						err = selfupdate.UpdateTo(latestAssetURL, exe)
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

// fileLabel is a small helper type that embeds a widget.Label and supports double-tap
// behavior so we can open a file when the user double-clicks it.
type fileLabel struct {
	widget.Label
	path     string
	onDouble func(string)
}

func (f *fileLabel) DoubleTapped(*fyne.PointEvent) {
	if f.onDouble != nil {
		f.onDouble(f.path)
	}
}

// openPath opens a file or folder using the platform's default handler.
func openPath(p string) {
	if p == "" {
		return
	}
	switch runtime.GOOS {
	case "darwin":
		_ = exec.Command("open", p).Start()
	case "windows":
		_ = exec.Command("cmd", "/C", "start", "", p).Start()
	default:
		_ = exec.Command("xdg-open", p).Start()
	}
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
			outputs, err := hotsheet.CreateHotsheet(inv, po, outdir)
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

			// helper state: selected index for Open Folder action
			var selectedIndex int = -1

			// Create a list whose items are of type fileLabel so we can respond to double-clicks.
			list := widget.NewList(
				func() int { return len(outputs) },
				func() fyne.CanvasObject {
					fl := &fileLabel{}
					fl.ExtendBaseWidget(fl)
					return fl
				},
				func(i widget.ListItemID, o fyne.CanvasObject) {
					fl := o.(*fileLabel)
					fl.path = outputs[i]
					fl.SetText(outputs[i])
					fl.onDouble = func(p string) { openPath(p) }
				},
			)

			// track selection so Open Folder knows which file's folder to open
			list.OnSelected = func(id widget.ListItemID) {
				selectedIndex = int(id)
			}
			list.OnUnselected = func(id widget.ListItemID) {
				selectedIndex = -1
			}

			// If there are no outputs, show a message; otherwise put the label in the top border
			// and let the list fill the center so it expands to available space.
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

			openFolderBtn := widget.NewButton("Open Folder", func() {
				if selectedIndex >= 0 && selectedIndex < len(outputs) {
					dir := filepath.Dir(outputs[selectedIndex])
					openPath(dir)
				} else if len(outputs) > 0 {
					// fallback to first output's folder
					dir := filepath.Dir(outputs[0])
					openPath(dir)
				}
			})

			// Place buttons at the bottom with spacer so they are right-aligned
			buttons := container.NewHBox(layout.NewSpacer(), openFolderBtn, widget.NewLabel("   "), doneBtn)

			// Use a border so the buttons stay at the bottom and the content fills the middle area.
			outWin.SetContent(container.NewBorder(nil, buttons, nil, nil, content))
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
