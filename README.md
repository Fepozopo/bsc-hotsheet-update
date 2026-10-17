# Hotsheet Updater

## Description

Hotsheet Updater is a small Go GUI application built with `nucular` that generates unified Excel hotsheets from a Sage 100 "Item Listing With Sales History" inventory report, an optional PO report, an optional sales history report, and an issue history report when sales history is supplied. For each product line found in the inventory report the app produces a single hotsheet file starting with a `Dashboard` sheet, followed by an `All Products` sheet, a `YTD Stock Priority` sheet, plus `Data Insights` and `Best Sellers` sheets and, when sales history is supplied, `Monthly History` and `MTO` sheets; it includes per-PO details when available and computes MTO (months-till-out) metrics. The `Data Insights` sheet shows a grouped sales outlook: `Counter Cards` by occasion on the left and `Other Products` by class and occasion on the right. It uses optional monthly dollar-sales history for completed-month, recency-weighted full-year projections rather than holiday selling windows.

## Motivation

This tool automates the manual work of assembling hotsheets from inventory and PO reports, reducing errors and saving time.

## Screenshots

### Legacy Data Insights sheet

This screenshot predates the overhaul. Current output replaces seasonal sections and status text with the grouped sales outlook described below, including PY Sales to Date when history is supplied.

![Legacy Data Insights sheet](images/data_insights.png)

## Requirements

- Go `1.26.2`
- `nucular` and the other Go module dependencies in `go.mod`
- Internet access is optional. The app checks for updates on startup when it can reach the public GitHub releases API, but it remains usable if the update check fails or there is no network connection.
- Native file pickers are launched through platform-native dialog systems:
  - macOS: `osascript`
  - Windows: native Common Item Dialog / shell APIs (no PowerShell dependency)
  - Linux: `zenity` or `kdialog`

## Quick Start

1. Clone the repository and open the project root.
2. Build for your platform using one of the `Makefile` targets. Example targets:
   - `make windows-amd64`
   - `make windows-arm64`
   - `make linux-amd64`
   - `make linux-arm64`
   - `make darwin-arm64` (macOS ARM64)
   - `make all`
   - `make clean` (removes `bin`)

Built binaries are written to the `bin/` directory.

## Usage (GUI)

1. Run the binary. The main window titled `Hotsheet Generator` opens.
2. Fill in:
   - Inventory Report (required): path to the inventory XLSX produced by Sage 100.
   - PO Report (optional): path to the PO XLSX (if omitted per-PO columns are not written).
   - Sales History (optional): path to the `IM_SalesHistory` XLSX (if omitted the `Monthly History` sheet is not written).
   - Issue History (optional, available only with Sales History): path to the `IM_IssueHistory` XLSX. BSC `Quantity Issued` is added to monthly `Quantity Sold`; negative issued months count as zero.
   - Best Sellers month range (optional, available after selecting Sales History): check the box and select From and To months and years. Both endpoint months are included for sales and issues. Without a checked range, Best Sellers uses inventory YTD Quantity Shipped (YTD sold plus nonnegative YTD issued) even when both history reports are selected.
   - Output Directory (optional): where generated files will be written (defaults to the current working directory).
3. Click `Generate Hotsheets`. The app validates inputs, shows a modal progress popup with a determinate progress bar, and performs the generation.
4. On success a `Created Hotsheets` modal popup lists generated files. Double-click an entry to open it, or use the Up/Down arrow keys to move through the list and press `Enter` to open the selected file. Hold `Option` on macOS (or `Alt` on other platforms) and press the bracketed letter in `Open Folder` or `Done` to open the selected file's folder or dismiss the popup. Press `Esc` to close the popup.
5. Throughout the main window, hold `Option` on macOS (or `Alt` on other platforms) and press the bracketed letter in the relevant label or button. The main form uses `I` for inventory report browsing, `P` for PO report browsing, `H` for sales history browsing, `S` for issue history browsing (when available), `O` for output directory browsing, `G` for generating hotsheets, `U` for checking for updates, and `Q` for quitting. On Windows the browse actions use the native Explorer-style Common Item Dialog instead of launching PowerShell.
6. When an update is available, hold `Option` on macOS (or `Alt` on other platforms) and press the bracketed `U` in `Update` or the bracketed `C` in `Continue`. Press `Esc` to close the popup as well. If you manually check for updates and you are already on the latest version, press the bracketed `O` in `OK` to dismiss the confirmation popup.

Behavior notes

- Inventory report is required; PO and sales history reports are optional. Issue history is optional only with sales history; providing issues without sales is rejected. When no PO report is supplied the output omits PO columns. When no sales history report is supplied the output omits `Monthly History` and `MTO`.
- The PO parser captures up to two PO lines per SKU; additional quantities are accumulated into the first PO slot.
- PO-only SKUs (SKUs present in PO but not in inventory) are skipped to avoid creating `UNKNOWN` product-line files.
- Output file naming: `{ProductLine}_hotsheet_YYYYMMDD.xlsx` (for example, `BAS_hotsheet_20260423.xlsx`).
- `Dashboard` is the first and initially active sheet. It has independently filtered top-15 donuts for inventory YTD quantity shipped (sold plus nonnegative issued) and inventory YTD dollars sold, plus a sales comparison chart. Each chart has its own class dropdown, defaulting to `All`, using inventory class descriptions. Rankings include all statuses, break ties by item code, and omit nonpositive values because donut slices cannot represent negatives; percentages describe only the selected top 15. These charts always use inventory YTD, regardless of the Best Sellers date range. Without monthly history, the comparison uses inventory YTD dollars versus full previous-year dollars. With monthly history, both comparison bars sum BSC Dollars Sold for completed months in the sales report's year and the same months in the prior year (a September run compares January–August; a January run has no completed months and shows zero). History also enables a clustered January–December monthly dollar-sales chart, with adjacent current/prior-year bars; the current month may be partial and future months normally zero. Missing history rows contribute zero, as in Monthly History. Class changes recalculate standard Excel formulas; no macros or manual refresh are required. Hidden `Dashboard Data` stores chart summaries. Use Excel with automatic calculation enabled for interactive filtering.
- Each output file contains `Dashboard`, `All Products`, `YTD Stock Priority`, `Data Insights`, and `Best Sellers`, plus `Monthly History` and `MTO` when the optional sales history report is supplied. `All Products` includes every inventory item (including Rundown and Discontinued) in source order. `YTD Stock Priority` excludes Rundown and Discontinued products and sorts the rest by ascending MTO YTD (lowest months till out first). Both inventory tabs share the former seasonal layout and have `QTY Shipped YTD` and `QTY Shipped PY` columns (inventory sold plus nonnegative inventory issued), as well as a filterable `Season` column before `Occasion`, showing Everyday, Winter, or Spring according to the same occasion mapping used for the former seasonal tabs. The optional `MTO` forecast includes issued quantities when Issue History is supplied.
- `Best Sellers` lists all inventory items in descending order of `Quantity Shipped` (SKU breaks ties). It includes item code, description, quantity shipped, dollars sold, quantity on hand, quantity committed (on sales orders plus on back order), quantity on PO, royalty code, class description, occasion, foil status, and status. Inventory quantities come from the current inventory report even when a sales month range is selected. Without a selected month range, it uses inventory YTD sold plus nonnegative inventory YTD issued and Dollars Sold YTD; with a range, it sums BSC sales-history Quantity Sold plus nonnegative issue-history Quantity Issued for the selected inclusive months, including across years, while Dollars Sold still comes only from sales history. The column always reads `Quantity Shipped`, even without issue history. Zero-shipment items remain at the bottom. Its header, status shading, and column filters match the other report sheets.
- `MTO` uses the sales-history run date as its snapshot date. It includes all undated POs in available quantity (`on hand + on PO - committed`, where committed is sales orders plus back orders) and forecasts BSC shipped units (sales plus nonnegative issued units when provided), not returned or transferred units. Recent completed observations of each calendar month receive higher weights (up to three years); completed zeros after the first positive shipment count, while pre-sale and partial current months do not. Forecast Demand is the next 12 calendar months' projected units; stockout and MTO are simulated month by month for up to 24 months. Partial months assume uniform daily sales. With at least two completed months since the first positive shipment, calendar months without an observation use the average of all completed months (including zero-sale months); months with observations retain their recent-year weighted forecasts. With fewer than two completed months, the forecast shows `Insufficient history`. The sheet lists only non-Rundown, non-Discontinued inventory items, puts the earliest stockouts first, and has column filters and the standard header color. Its Class Description uses the standard sheets' SKU-prefix rules, and its filterable Season column immediately before Occasion uses the same Everyday/Winter/Spring mapping as the standard sheets. Numeric MTO values use the lighter `MTO YTD` colors (red at ≤1 month, yellow through 3, green above 3); nonnumeric values such as `Insufficient history` remain uncolored. History Coverage shows completed observed months and years from the first positive shipment, not a verified launch date. To test an additional order without changing baseline MTO, enter a nonnegative whole number in the yellow `Proposed PO Units` cell. `MTO with Proposed PO` and `Stockout Month with Proposed PO` recalculate in Excel from a hidden month-by-month forecast; blank means zero additional units, and results past the 24-month horizon show `>24` / `Not within 24 months`. Existing POs are already counted in Available Quantity. Both actual and proposed POs are assumed available immediately, which can make projected stockout optimistic.
- For product line `2021` only, `MTO` excludes history before February 2026 for these SKUs: `BD1001FJ`, `BD1006F`, `BD1028`, `BD1044`, `BD1046F`, `BD1048FJ`, `BP1002`, `CS1003`, `CS1005`, `FC1005`, `GR1016FJ`, `HY1048FB`, `HY1071`, `HY1071B`, `LV1010`, `MD1011`, `MI1011`, `MI1014`, `PJ1011`, `PJ1031`, `PJ1032`, `RL1003`, `SP1001F`, `TY1007`, `TY1019`, `TY1019B`, `TY1020`, `TY1023`, `TY1023B`, `TY1025`, `TY1025B`, `TY1028F`, and `VD1029J`. The cutoff applies to sold and issued quantities, History Coverage, baseline forecasts, and proposed-PO scenarios. Existing insufficient-history rules still apply if fewer than two eligible completed months remain. All other SKUs, product lines, and sheets keep their existing behavior and full source history.
- `Monthly History` lists BSC-warehouse sales and optional nonnegative issued values only: one row per item/year/metric, January–December in separate columns, and the inventory item's status in the last column. The sheet has filters on all columns and the same header and status shading as the standard sheets. Inventory YTD/PY values remain unchanged; Data Insights projections use optional monthly dollar-sales history.
- `Data Insights` has two side-by-side areas: `Counter Cards` grouped by occasion and `Other Products` grouped by class and occasion, alphabetically ordered without seasonal sections or status text. Columns are `Occasion`, `Holiday / Date`, inventory `YTD Sales`, optional `PY Sales to Date`, inventory full-year `PY Sales`, and `Projected YoY`. Fixed holidays use stable dates (for example, `Feb 14`); movable holidays use names (for example, `Easter`), and Graduation and undated occasions show `N/A`.
- With monthly history, Data Insights uses the sales report's run date, never today's date. PY Sales to Date sums prior-year dollars for January through the last completed current-year month. Full-year projected sales equal observed completed current-year dollars plus monthly estimates for the entire current month through December; incomplete current-month sales are not extrapolated or added again. Displayed inventory YTD can therefore differ from completed-month actuals used in the forecast.
- Data Insights aggregates BSC `Dollars Sold` at the combined-row level before forecasting. It excludes months before the combined row's first positive dollar sale, counts subsequent supplied zero months, and requires at least two distinct completed year/month observations. A missing SKU/year record is unknown, not another observed zero. For each calendar month, at most three recent observations receive MTO-style 3:2:1 weights (2:1 with two years). Months without comparable history use the average of the retained observed months; incomplete calendar coverage is marked `limited-history estimate`. Signed dollar corrections are retained. Without sufficient history, including when no report is supplied, Projected YoY shows `Insufficient history`. Without a nonzero full-year inventory PY baseline it instead shows `N/A — no PY baseline`. Each table total independently forecasts its combined monthly dollar history using the same completed-month, first-positive-sale, weighting, and two-month minimum rules, then compares that forecast with the table's full-year inventory PY total. An empty or unprojectable occasion does not block a total with sufficient combined history. Total forecasts can differ from the sum of individual occasion forecasts because aggregation changes history coverage and fallback estimates; missing records still remain unknown.

## Logs

The application writes JSON-formatted logs into a `logs-bsc` directory inside the OS temporary directory (`os.TempDir()`). Filenames include a timestamp and the logical logger name, with optional product/occasion suffixes. Example patterns produced by the logger:

- `2006-01-02_150405.000000000_name.log`
- `2006-01-02_150405.000000000_name-product-occasion.log`

Logger implementation: `helpers/slog_logger.go`. Callers must close the returned `io.Closer` to flush buffered entries (the code already defers `Close()`).

## Auto-update

On startup the GUI checks the public GitHub releases API for the latest version.

- If a newer release is detected, the app prompts the user and offers both `Update` and `Continue`.
- If the user chooses `Update`, the app downloads the release asset, replaces the running executable, and restarts the new binary.
- If the user chooses `Continue`, the current version remains usable and the app continues normally.
- If the update check fails (for example, no internet access), the app continues running and shows a non-blocking status message instead of preventing use.
- If an update attempt fails, the app shows an error popup but the current version remains usable.

## Implementation details

- Entry point: `main.go` sets up logging and launches the Nucular GUI via `internal/gui`.
- GUI: `internal/gui/app.go`, `internal/gui/state.go`, `internal/gui/actions.go`, `internal/gui/render_main.go`, and `internal/gui/render_popups.go` contain the immediate-mode UI, popups, input handling, determinate generation-progress display, and background-task coordination.
- Native dialogs and file opening: `internal/gui/native_dialogs.go`, `internal/gui/native_dialogs_nonwindows.go`, `internal/gui/native_dialogs_windows.go`, `internal/gui/open.go`, and `internal/gui/open_windows.go` preserve native file pickers and platform-specific open behavior. Windows now uses the Common Item Dialog for both file and folder browsing and `ShellExecuteW` for opening files/folders without spawning a terminal window.
- Auto-update transport: `internal/update/service.go` checks the public GitHub releases API, selects the correct release asset for the active platform, applies updates, and restarts the executable.
- Hotsheet generation: `hotsheet/generate.go` exposes `hotsheet.Generate(...)`, accepts an optional progress callback for coarse determinate progress updates, and orchestrates the report pipeline. The package is now split by responsibility: `hotsheet/inventory_reader.go` parses the inventory export, `hotsheet/po_reader.go` merges optional PO data, `hotsheet/sales_history_reader.go` streams optional BSC monthly sales and issues into the item model, `hotsheet/monthly_history_sheet.go` writes the optional history sheet, `hotsheet/best_sellers.go` ranks the inventory and writes Best Sellers, `hotsheet/mto_forecast.go` models seasonal stockout and `hotsheet/mto_sheet.go` writes the MTO tab, `hotsheet/product_line.go` groups entries by product line, `hotsheet/standard_sheets.go` writes the `All Products` and `YTD Stock Priority` tabs, `hotsheet/data_insights_sheet.go` renders the `Data Insights` worksheet, `hotsheet/data_insights_rows.go` builds grouped Data Insights rows, `hotsheet/data_insights_projection.go` contains stable holiday labels and completed-month dollar projection logic, `hotsheet/workbook.go` creates and saves workbooks, `hotsheet/styles.go` centralizes workbook styles, and `hotsheet/parsing.go`, `hotsheet/occasion.go`, and `hotsheet/entry.go` hold shared parsing, occasion mapping, and core model definitions.
- Logging: `helpers/slog_logger.go` creates buffered JSON writers into `logs-bsc` under the system temp directory.
- Version: `internal/version/version.go`.
- Build: `Makefile` provides cross-compile targets and passes explicit `nucular` backend tags per platform.

## Troubleshooting

- Auto-update failed: the app should still remain usable. Ensure internet connectivity if you want update checks/downloads to succeed, and ensure the app has permission to replace the executable if you choose to update.
- No logs: check your OS temp directory for a `logs-bsc` folder and file permissions.
- If a browse button fails to open a picker, ensure your platform dialog support is available (`osascript` on macOS, built-in Windows shell/Common Item Dialog support on Windows, `zenity` or `kdialog` on Linux).
- If you build manually without the expected `nucular` tag for your platform, the GUI may not use the intended backend. Prefer the provided `Makefile` or the explicit `go run -tags ...` examples above.
