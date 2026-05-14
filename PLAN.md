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

## Proposed function responsibilities

### Inventory loading helpers

These helpers should be responsible for:

- opening the inventory workbook
- reading the inventory sheet rows
- parsing inventory rows into `entry` values
- storing parsed entries in `invMap`

### PO loading helpers

These helpers should be responsible for:

- opening the optional PO workbook
- reading the PO sheet rows
- applying PO quantities to matching inventory entries
- skipping PO-only SKUs as today

### Row materialization helpers

These helpers should be responsible for:

- converting enriched entries into the reporting-ready slice used downstream
- preserving existing derived values and class-prefix rules
- keeping the workbook-writing layer separate from parsing

### Grouping helpers

These helpers should be responsible for:

- grouping entries by product line
- preparing each product line’s data set for workbook generation
- keeping the grouping logic isolated from file I/O and formatting

### Workbook helpers

These helpers should be responsible for:

- creating a new workbook for a product line
- writing the standard sheets
- writing the `Data Insights` sheet
- applying styles, widths, filters, and comments
- saving the workbook to disk

## Suggested decomposition targets

The refactor will likely benefit from extracting functions similar to these:

- `loadInventoryEntries(...)`
- `mergePOData(...)`
- `buildProductLineGroups(...)`
- `buildProductLineWorkbook(...)`
- `writeStandardSheets(...)`
- `writeDataInsightsSheet(...)` already exists and should be reused as-is if possible
- `saveWorkbook(...)`

The exact names can change, but the responsibilities should stay separated.

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
