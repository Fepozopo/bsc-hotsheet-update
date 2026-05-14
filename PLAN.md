# PLAN

## Goal

Refactor `CreateFromReports()` into smaller, focused functions so the workbook generation flow is easier to maintain and easier to reuse in future features.

This refactor should preserve the current behavior of the application, including:

- inventory parsing
- optional PO parsing
- product-line workbook generation
- existing sheet output, including `Data Insights`
- current workbook styling and file naming behavior

## Why this refactor matters

`CreateFromReports()` currently does too many things in one place:

- opens and reads source workbooks
- parses inventory rows into entries
- merges PO data
- groups rows by product line
- computes derived values
- builds workbook sheets
- applies styles and formatting
- writes files to disk

Breaking this into smaller functions should make the code:

- easier to read
- easier to test
- easier to reuse in future reports
- safer to modify without accidentally affecting unrelated behavior

## Desired structure

The refactor should separate the work into clear stages, roughly like this:

1. Load inventory rows
2. Build the inventory entry map
3. Load and merge PO rows when provided
4. Materialize report-ready entries
5. Group entries by product line
6. Build a workbook for each product line
7. Write all sheets and save the file

## Concrete function breakdown

This is the recommended helper breakdown for the refactor. The exact names can change, but the responsibilities should stay close to this split.

### 1. Top-level orchestration

#### `CreateFromReports(inventoryPath, poPath, outputDir string) ([]string, error)`

Keep this function as the orchestration layer only. It should:

- create the logger
- load inventory entries
- merge PO data when provided
- group entries by product line
- loop through each product line
- call the workbook builder for each group
- collect output file paths
- return the final list of generated files

It should no longer contain the detailed parsing or workbook-writing logic directly.

---

### 2. Inventory loading

#### `loadInventoryEntries(inventoryPath string, logger *slog.Logger) (map[string]*entry, error)`

Responsibilities:

- open the inventory workbook
- read the inventory sheet rows
- parse each inventory item into an `entry`
- populate `invMap` by SKU
- preserve the current empty-row and run-date stop logic
- preserve existing inventory parsing behavior

This helper should return the populated inventory map only.

#### `parseInventoryEntry(rows [][]string, rowNum int, logger *slog.Logger) (*entry, bool)`

Responsibilities:

- parse one inventory item from the inventory worksheet rows
- return the parsed `entry`
- return a boolean indicating whether the row should stop parsing or be skipped

This keeps the row-parsing logic separate from file opening and loop control.

---

### 3. PO merging

#### `mergePOData(poPath string, invMap map[string]*entry, logger *slog.Logger) error`

Responsibilities:

- open the optional PO workbook
- read PO rows
- match PO rows to inventory entries by SKU
- apply PO quantities to matching entries
- preserve the current rule that skips PO-only SKUs
- preserve the current behavior of collecting up to two PO lines per SKU

This helper should only modify the provided `invMap`.

#### `applyPOToEntry(e *entry, poRow []string, ...)`

Optional smaller helper if the PO logic still feels too dense after extraction.

Responsibilities:

- normalize the PO number
- choose the correct quantity column based on status
- assign the PO to the first available slot on the entry

---

### 4. Product-line grouping

#### `groupEntriesByProductLine(invMap map[string]*entry, logger *slog.Logger) map[string][]*entry`

Responsibilities:

- iterate the populated inventory map
- skip entries with empty product lines when appropriate
- group entries by `ProductLine`
- return a map of product line to entries

This should stay focused on grouping only.

#### `sortEntriesForProductLine(entries []*entry)`

Responsibilities:

- sort entries into a stable order before workbook generation
- preserve the current ordering behavior used for each product line

If sorting rules become more complex later, this can be expanded or split further.

---

### 5. Product-line workbook builder

#### `buildProductLineWorkbook(productLine string, entries []*entry, outputDir, dateStr string, logger *slog.Logger) (string, error)`

Responsibilities:

- create the workbook for one product line
- write all standard report sheets
- write the `Data Insights` sheet
- apply sheet-level widths, filters, and comments
- save the workbook to disk
- return the output file path

This is the main workbook-level helper that `CreateFromReports()` should call for each product line.

#### `newProductLineWorkbook() *excelize.File`

Optional helper if workbook initialization needs to be isolated.

Responsibilities:

- create the workbook
- create the standard sheets
- delete the default sheet
- set the active sheet

---

### 6. Standard sheet writing

#### `writeStandardSheets(f *excelize.File, entries []*entry, hasPO bool) error`

Responsibilities:

- write the existing `Everyday`, `Winter`, and `Spring` sheets
- write headers
- write data rows
- apply MTO calculations
- apply row styles
- preserve current workbook behavior for the standard sheets

This should be the main helper for the existing three worksheet tabs.

#### `writeStandardSheetHeaders(f *excelize.File, sheetName string, headers []string, hasPO bool) error`

Responsibilities:

- write the header row for one sheet
- apply the current header formatting
- add MTO header comments where needed

#### `writeStandardSheetRows(f *excelize.File, sheetName string, entries []*entry, hasPO bool, monthsThrough float64) error`

Responsibilities:

- compute row values for each entry
- calculate MTO values
- write row values
- apply per-row formatting
- preserve current class-prefix logic and status coloring behavior

#### `applyStandardSheetWidths(f *excelize.File, hasPO bool) error`

Responsibilities:

- set column widths for the standard sheets
- keep the current readability settings consistent

#### `applyStandardSheetFilters(f *excelize.File) error`

Responsibilities:

- apply autofilter to the standard sheets
- keep the current filter range and behavior the same

---

### 7. `Data Insights` sheet reuse

#### `writeDataInsightsSheet(f *excelize.File, entries []*entry) error`

This helper already exists and should remain the dedicated builder for the `Data Insights` tab.

In the refactor, `buildProductLineWorkbook(...)` should call it rather than duplicating any of its logic.

---

### 8. File output

#### `saveWorkbook(f *excelize.File, outputDir, productLine, dateStr string) (string, error)`

Responsibilities:

- build the final output file path
- sanitize the product line name
- save the workbook
- return the generated path

This keeps file naming and disk output separate from workbook construction.

## Proposed call flow

The intended flow after the refactor should look like this:

1. `CreateFromReports()`
2. `loadInventoryEntries()`
3. `mergePOData()` if a PO report was supplied
4. `groupEntriesByProductLine()`
5. for each product line:
   - `sortEntriesForProductLine()`
   - `buildProductLineWorkbook()`
     - `writeStandardSheets()`
       - `writeStandardSheetHeaders()`
       - `writeStandardSheetRows()`
       - `applyStandardSheetWidths()`
       - `applyStandardSheetFilters()`
     - `writeDataInsightsSheet()`
     - `saveWorkbook()`

## Step-by-step implementation order

Use this order to keep the refactor safe and easy to verify:

### Step 1: Extract inventory loading first

Start by moving the inventory workbook open/read/parse work into `loadInventoryEntries()` and `parseInventoryEntry()`.

Why first:

- it is the foundation for everything else
- it does not change workbook output yet
- it makes later extraction steps easier because the parsed inventory data is isolated

### Step 2: Extract PO merging next

Move the optional PO workbook parsing and merge logic into `mergePOData()` and, if helpful, `applyPOToEntry()`.

Why second:

- it builds directly on the inventory map from Step 1
- it keeps PO behavior isolated before workbook generation is touched
- it reduces the amount of logic left inside `CreateFromReports()` early

### Step 3: Extract product-line grouping

Move product-line grouping into `groupEntriesByProductLine()` and any stable sorting into `sortEntriesForProductLine()`.

Why third:

- grouping is independent of workbook styling
- it defines the unit of work for the workbook builder
- it makes the next step cleaner because each group is already prepared

### Step 4: Extract workbook creation per product line

Create `buildProductLineWorkbook()` and, if needed, `newProductLineWorkbook()` so workbook setup is isolated from orchestration.

Why fourth:

- it establishes a single place to build and save a workbook
- it sets the stage for moving sheet-writing logic out of `CreateFromReports()`
- it makes the top-level function much smaller once wired in

### Step 5: Extract standard sheet writing

Move the existing Everyday, Winter, and Spring sheet writing logic into `writeStandardSheets()` and smaller helpers such as:

- `writeStandardSheetHeaders()`
- `writeStandardSheetRows()`
- `applyStandardSheetWidths()`
- `applyStandardSheetFilters()`

Why fifth:

- the standard sheets are the largest remaining body of workbook logic
- splitting them after workbook setup keeps the new structure easier to follow
- this step should preserve the current sheet output while reducing duplication risk

### Step 6: Reuse the existing `Data Insights` builder

Keep `writeDataInsightsSheet()` as the dedicated helper and call it from the workbook builder.

Why sixth:

- it already exists and should not be reworked unless needed
- it should fit naturally once workbook creation is isolated
- it avoids mixing new refactor work with an already working helper

### Step 7: Reduce `CreateFromReports()` to orchestration only

Once the helpers exist, trim `CreateFromReports()` down to the pipeline that wires the helpers together.

Why seventh:

- by this point the lower-level behavior is already extracted
- the remaining function should be easy to read and review
- any final cleanup here is mainly about clarity, not behavior

### Step 8: Verify after each major extraction

Run validation as the refactor progresses, not only at the end.

Recommended checkpoints:

- after inventory extraction
- after PO extraction
- after workbook builder extraction
- after standard sheet extraction
- after the final orchestration cleanup

## What should stay the same

The refactor should not change the observable behavior unless it is clearly part of cleanup:

- workbook output should still be generated per product line
- existing sheet contents should remain the same
- `Data Insights` should still be included
- PO-only SKUs should still be skipped
- file naming should stay the same
- style and layout behavior should remain compatible with current output

## Proposed implementation approach

### Phase 1: Map the current flow

- Identify the major sections inside `CreateFromReports()`
- Group related logic into cohesive responsibilities
- Identify what data needs to move between helper functions

### Phase 2: Extract parsing helpers

- Move inventory parsing into a dedicated helper
- Move PO parsing into a dedicated helper
- Keep the same parsing behavior while reducing function size

### Phase 3: Extract workbook-building helpers

- Move sheet creation and styling into workbook-specific helpers
- Keep workbook generation separate from data parsing
- Reuse the existing `Data Insights` sheet builder during the split

### Phase 4: Preserve behavior

- Confirm generated files still match current expectations
- Confirm the refactor does not change sheet order, file names, or key calculations
- Confirm error handling remains clear and actionable

### Phase 5: Optional cleanup

- If common workbook formatting logic emerges, extract shared helpers
- If repeated style setup appears, centralize it
- If data transformations repeat, isolate them into reusable functions

## Potential helper categories

### Parsing helpers

- read workbook rows
- parse inventory row values
- parse PO row values
- normalize text and numbers

### Transformation helpers

- merge PO values into entries
- compute derived values
- group entries by product line
- prepare report rows for sheet generation

### Workbook helpers

- create workbook and sheets
- write headers
- apply styles
- set widths and filters
- save output file

## Acceptance criteria

- `CreateFromReports()` is reduced to an orchestration function
- inventory parsing is split into smaller helpers
- PO parsing is split into smaller helpers
- workbook creation is split into smaller helpers
- `Data Insights` remains supported
- the refactor preserves current output behavior
- the refactor makes future reuse of parsing or workbook logic easier

## Concrete build checklist

### Flow decomposition

- [ ] Identify the distinct responsibilities currently handled inside `CreateFromReports()`
- [ ] Define the helper boundaries before editing code
- [ ] Decide which parts should become parsing helpers versus workbook helpers
- [ ] Confirm the existing `Data Insights` builder can be reused without redesign

### Inventory parsing extraction

- [ ] Move inventory workbook opening and row reading into a helper
- [ ] Move inventory row parsing into a helper that returns `entry` values
- [ ] Keep the current inventory parsing behavior unchanged
- [ ] Preserve logging behavior around inventory parsing

### PO parsing extraction

- [ ] Move optional PO workbook opening and row reading into a helper
- [ ] Move PO matching and merge logic into a helper
- [ ] Preserve the existing PO-only SKU skip behavior
- [ ] Preserve the current per-SKU PO assignment behavior

### Grouping and preparation extraction

- [ ] Move product-line grouping into a helper
- [ ] Move report-ready row preparation into a helper
- [ ] Keep the current derived-value calculations intact
- [ ] Preserve the current class-prefix behavior

### Workbook generation extraction

- [ ] Extract workbook creation into a helper
- [ ] Extract standard sheet writing into a helper
- [ ] Keep `Data Insights` writing reusable from the workbook layer
- [ ] Keep styles, widths, filters, and comments in workbook-specific helpers

### Orchestration cleanup

- [ ] Reduce `CreateFromReports()` to a readable top-level pipeline
- [ ] Keep error handling meaningful at each stage
- [ ] Keep the function easy to scan from top to bottom
- [ ] Avoid changing behavior while restructuring

### Verification

- [ ] Confirm generated files still build successfully
- [ ] Confirm the output workbook structure is unchanged
- [ ] Confirm `Data Insights` still appears in each generated workbook
- [ ] Confirm PO parsing still behaves the same
- [ ] Confirm file naming and save behavior are unchanged
- [ ] Confirm the refactor improves readability without adding unnecessary complexity

## Status

Planning only. `CreateFromReports()` has not been refactored yet.
