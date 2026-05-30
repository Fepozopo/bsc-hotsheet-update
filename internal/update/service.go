package update

import (
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"os/exec"
	"runtime"
	"strings"

	"github.com/blang/semver"
	"github.com/rhysd/go-github-selfupdate/selfupdate"
)

type CheckResult struct {
	CurrentVersion  semver.Version
	LatestVersion   semver.Version
	AssetURL        string
	UpdateAvailable bool
}

type githubRelease struct {
	TagName string               `json:"tag_name"`
	Assets  []githubReleaseAsset `json:"assets"`
}

type githubReleaseAsset struct {
	Name               string `json:"name"`
	BrowserDownloadURL string `json:"browser_download_url"`
}

// CurrentReleaseAssetName returns the GitHub release asset name that matches the current
// platform so the auto-updater can find the correct binary for this build.
func CurrentReleaseAssetName() string {
	base := "hotsheet"
	switch runtime.GOOS {
	case "windows":
		return fmt.Sprintf("%s-%s-%s.exe", base, runtime.GOOS, runtime.GOARCH)
	default:
		return fmt.Sprintf("%s-%s-%s", base, runtime.GOOS, runtime.GOARCH)
	}
}

// DetectLatestRelease queries the GitHub releases API, extracts the version tag, and
// returns the download URL for the asset that matches the current platform.
func DetectLatestRelease(repo string) (semver.Version, string, error) {
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

	expectedAsset := CurrentReleaseAssetName()
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

// CheckForUpdates resolves the latest release and determines whether it is newer than the
// provided current version string.
func CheckForUpdates(repo, currentVersion string) (CheckResult, error) {
	currentVer, err := semver.Parse(currentVersion)
	if err != nil {
		return CheckResult{}, fmt.Errorf("could not parse current version %q: %w", currentVersion, err)
	}

	latestVersion, assetURL, err := DetectLatestRelease(repo)
	if err != nil {
		return CheckResult{}, err
	}

	return CheckResult{
		CurrentVersion:  currentVer,
		LatestVersion:   latestVersion,
		AssetURL:        assetURL,
		UpdateAvailable: latestVersion.GT(currentVer),
	}, nil
}

// ApplyUpdate downloads the asset at assetURL and replaces the executable at exePath.
func ApplyUpdate(assetURL, exePath string) error {
	return selfupdate.UpdateTo(assetURL, exePath)
}

// RestartExecutable launches the updated executable with the supplied arguments.
func RestartExecutable(exe string, args []string) error {
	cmd := exec.Command(exe, args...)
	cmd.Env = os.Environ()
	return cmd.Start()
}
