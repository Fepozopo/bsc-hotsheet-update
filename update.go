package main

import (
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"os/exec"
	"runtime"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	"github.com/Fepozopo/bsc-hotsheet-update/internal/version"
	"github.com/blang/semver"
	"github.com/rhysd/go-github-selfupdate/selfupdate"
)

type githubRelease struct {
	TagName string               `json:"tag_name"`
	Assets  []githubReleaseAsset `json:"assets"`
}

type githubReleaseAsset struct {
	Name               string `json:"name"`
	BrowserDownloadURL string `json:"browser_download_url"`
}

// currentReleaseAssetName returns the GitHub release asset name that matches the current
// platform so the auto-updater can find the correct binary for this build.
func currentReleaseAssetName() string {
	base := "hotsheet"
	switch runtime.GOOS {
	case "windows":
		return fmt.Sprintf("%s-%s-%s.exe", base, runtime.GOOS, runtime.GOARCH)
	default:
		return fmt.Sprintf("%s-%s-%s", base, runtime.GOOS, runtime.GOARCH)
	}
}

// detectLatestRelease queries the GitHub releases API, extracts the version tag, and
// returns the download URL for the asset that matches the current platform.
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

// checkForUpdates runs the release check in the background so the UI stays responsive.
// If a newer release exists, the app forces the update path because the rest of the flow
// assumes the binary is kept current.
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
