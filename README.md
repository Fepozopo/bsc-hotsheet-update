# Hotsheet Updater

## Description

Hotsheet Updater is a small Go GUI (Fyne) application that generates unified Excel hotsheets from a Sage 100 "Item Listing With Sales History" inventory report and an optional PO report. For each product line found in the inventory report the app produces a single hotsheet file with three operational sheets (Everyday, Winter, Spring) plus a `Data Insights` summary sheet, includes per-PO details when available, and computes MTO (months-till-out) metrics.

## Motivation

This tool automates the manual work of assembling hotsheets from inventory and PO reports, reducing errors and saving time.

## Requirements

- Go 1.26.2
- Fyne v2.5.3 and other Go module dependencies
- A C compiler for CGO (Fyne requires C bindings):
  - The Makefile uses `zig` for many cross-compilation targets; edit targets to use `clang`/`gcc` if preferred.
  - The darwin (macOS) target uses `clang`.
- Internet access is required for the built-in auto-update check (uses GitHub releases).

## Quick Start

1. Clone the repository and open the project root.
2. Build for your platform using one of the Makefile targets. Example targets:
   - `make windows-amd64`
   - `make windows-arm64`
   - `make linux-amd64`
   - `make linux-arm64`
   - `make darwin-arm64` (macOS ARM64)
   - `make all`
   - `make clean` (removes `bin`)

Built binaries are written to the `bin/` directory. For local testing `go run .` launches the GUI directly.

## Usage (GUI)

1. Run the binary. The main window titled "Hotsheet Generator" opens.
2. Fill in:
   - Inventory Report (required): path to the inventory XLSX produced by Sage 100.
   - PO Report (optional): path to the PO XLSX (if omitted per-PO columns are not written).
   - Output Directory (optional): where generated files will be written (defaults to the current working directory).
3. Click "Generate Hotsheets". The app validates inputs, shows a progress dialog, and performs the generation.
4. On success a "Created Hotsheets" window lists generated files. Double-click an entry to open it, or use "Open Folder" to reveal the containing folder. Click "Done" to close and clear inputs to run again.

Behavior notes

- Inventory report is required; PO report is optional. When no PO report is supplied the output omits PO columns.
- The PO parser captures up to two PO lines per SKU; additional quantities are accumulated into the first PO slot.
- PO-only SKUs (SKUs present in PO but not in inventory) are skipped to avoid creating "UNKNOWN" product-line files.
- Output file naming: `{ProductLine}_hotsheet_YYYYMMDD.xlsx` (e.g., `BAS_hotsheet_20260423.xlsx`).
- Each output file contains four sheets: "Everyday", "Winter", "Spring", and "Data Insights". Header comments explain the MTO calculations.

## Logs

The application writes JSON-formatted logs into a `logs-bsc` directory inside the OS temporary directory (os.TempDir()). Filenames include a timestamp and the logical logger name, with optional product/occasion suffixes. Example patterns produced by the logger:

- `2006-01-02_150405.000000000_name.log`
- `2006-01-02_150405.000000000_name-product-occasion.log`

Logger implementation: `helpers/slog_logger.go`. Callers must close the returned io.Closer to flush buffered entries (the code already defers Close()).

## Auto-update

On startup the GUI checks GitHub releases (uses `rhysd/go-github-selfupdate`). If a newer release is detected the app prompts the user. If the update is accepted the app downloads the release asset, replaces the running executable, and restarts the new binary. If the update is declined the app will exit. Update errors are shown in an error dialog.

## Implementation details

- Entry point: `main.go` sets up logging and launches the GUI flow (`selectFiles`).
- GUI & update-checks: `app.go` contains the UI, input validation, progress dialogs, and the auto-update check.
- Hotsheet generation: `hotsheet/create_from_reports.go` parses the inventory and optional PO reports and writes XLSX files; `hotsheet/copy_hotsheet.go` contains a simple copy helper; `hotsheet/util.go` includes parsing and mapping helpers.
- Logging: `helpers/slog_logger.go` creates buffered JSON writers into `logs-bsc` under the system temp directory.
- Version: `internal/version/version.go` currently holds the app version (v2.1.1).
- Build: `Makefile` provides cross-compile targets (uses CGO and `zig` by default).

## Troubleshooting

- Auto-update failed: ensure internet connectivity and that the app has permission to replace the executable.
- No logs: check your OS temp directory for a `logs-bsc` folder and file permissions.
- Build failures due to missing `zig`: either install `zig` (recommended for cross-compiles) or edit the `Makefile` targets to use your local `clang`/`gcc` toolchain.
