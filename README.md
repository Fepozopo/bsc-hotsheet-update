# Hotsheet Updater

## Description

This is a Go program with a small GUI (Fyne) that updates Excel hotsheets using data from an "Item Listing With Sales History" report pulled from Sage 100. It can update multiple product lines and their respective sheets in one run and automatically matches and updates PO numbers and quantities for each SKU. The program shows a GUI for selecting files and a console progress indicator during updates. Detailed logs are written to a `logs-bsc` folder inside the system temporary directory.

## Motivation

This tool replaces the manual process of copying PO and inventory values into hotsheets. It significantly reduces time and human error by automating:

- Matching SKU rows,
- Updating PO quantities and numbers,
- Copying a source hotsheet and writing an updated file,
- Writing detailed logs for troubleshooting.

## Requirements

- Go 1.16+ (recommended latest stable Go)
- A C compiler for CGO (Fyne requires C bindings):
  - The Makefile uses `zig` for cross compilation targets. If you don't have zig, you can adjust the Makefile to use a local C compiler (e.g., `clang` or `gcc`).
  - For macOS builds the provided target uses `clang`.
- Fyne GUI dependencies are handled by Go modules in `go.mod`.
- Internet access is required for the built-in auto-update check to contact GitHub releases.

## Quick Start

1. Clone the repository and open the project root.
2. Build for your platform using one of the Makefile targets. Example targets:
   - `make windows-amd64`
   - `make windows-arm64`
   - `make linux-amd64`
   - `make linux-arm64`
   - `make darwin-arm64` (macOS ARM64)
   - `make all` (builds all targets)
   - `make clean` (removes the `bin` folder)

Built binaries are placed in the `bin` directory and are named like `hotsheet-windows-amd64.exe`, `hotsheet-linux-amd64`, `hotsheet-darwin-arm64`, etc.

Notes about building:

- The Makefile sets `CGO_ENABLED=1` and uses `zig cc` for cross compilation. If you prefer another toolchain, edit the corresponding target.

## Usage (GUI)

1. Run the built binary (or run `go run .` during development).
2. The application window asks which hotsheet you want to update. Options include:
   - All (updates 21c, BSC, SMD — BJP is excluded from the "All" option)
   - 21c
   - BJP
   - BSC
   - SMD
3. Hotsheet file selection:
   - If you choose a single product (e.g., `BSC`), only that product's hotsheet file selector is shown and is required.
   - If you choose `All`, three hotsheet file entries are required (in the UI these correspond to the 21c, BSC, and SMD hotsheet selectors). You must fill all three when selecting `All`.
   - The "Next" button (to select reports) is enabled only after the required hotsheet file(s) are chosen.
4. Reports:
   - Inventory report: required
   - PO report: required
   - BN report: optional — enable the "Include BN report (optional)" checkbox to show the BN selector (if checked, a BN file must be chosen).
5. Submit: After validation, the app copies the selected hotsheet(s) and performs updates using the selected reports. Progress is shown in the console and operations are logged.

Important behaviors:

- "All" maps to the following required hotsheet files: 21c, BSC, SMD (in that order). The program will attempt to update these three product lines in that order.
- If you want to update BJP, choose it explicitly (do not use "All" — BJP is not included in "All").

## Logs

Logs are written to a `logs-bsc` directory inside the OS temporary directory. The logger creates files with a timestamped name and includes product/occasion information when available. Example file name patterns:

- `YYYY-MM-DD_HHMMSS.sssssss_NAME.log`
- `YYYY-MM-DD_HHMMSS.sssssss_NAME-PRODUCT-OCCASION.log`

The logger implementation lives in `helpers/create_logger.go`.

## Auto-update

The GUI performs a GitHub release check (using `go-github-selfupdate`) on startup:

- If a newer release is found, the user is prompted to update.
- When an update is accepted the app downloads the new asset, attempts to replace the executable, and restarts the new binary.
- If update checks fail, an error dialog is shown.

## Example: Update struct usage

You can still configure which columns and sheets to update via the `Update{...}` struct usage in code. An example snippet used in the codebase (this is illustrative; edit code in `hotsheet` package where needed):
`Update{ Hotsheet: fileHotsheetNew, Sheet: "EVERYDAY", SkuCol: "C", OnHandCol: "D", ... }`

(See `hotsheet/update.go` and the various `case_*.go` files for concrete examples.)

## Files of interest

- `main.go` — app entry; orchestrates file selection and per-product updating.
- `app.go` — GUI logic, file selection, validation, and update checks.
- `hotsheet/` — core logic for copying and updating hotsheet Excel files:
  - `case_21c.go`, `case_bsc.go`, `case_bjp.go`, `case_smd.go`
  - `copy_hotsheet.go`, `update.go`
- `helpers/` — logger and progress bar helpers.
- `internal/version/version.go` — application version string (update before releases).
- `Makefile` — build targets and instructions.

## Troubleshooting

- If the GUI reports the auto-update failed, check internet connectivity and permissions to replace the executable.
- If logs are missing, check your OS temp directory for `logs-bsc` and verify the process has permission to write to it.
- If a build fails due to missing `zig`, either install `zig` (recommended for cross-compiles) or edit the Makefile to use another CC for your environment.
