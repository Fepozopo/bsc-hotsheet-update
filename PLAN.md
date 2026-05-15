# PLAN

## Goal

Update the `Data Insights` sheet so it includes a second grouped table for all non-counter-card products.

The sheet should show:

- a left-side table for **Counter Cards**
- a right-side table for **Other Products**
- both tables using the same column layout as the current occasion tables:
  - `Class`
  - `Date`
  - `YTD Sales`
  - `PY Sales`
  - `Projected YoY`

The `Data Insights` title should continue to span the full width of all tables at the top, and it must extend to cover the new right-side table. Add subtitles for both sides:

- left side: `Counter Cards`
- right side: `Other Products`

The new right-side table should start in **column I**.

## Why this change matters

Right now, `Data Insights` separates cards into occasions, which makes it hard to review the rest of the class descriptions in one place. Adding a separate table for non-counter-card products will make the sheet more useful for quick comparison and reporting.

This change should:

- keep the existing occasion-based behavior for counter cards
- add a dedicated grouped view for everything else
- preserve the current formatting style as much as possible
- keep the sheet readable by balancing the left and right sections

## Desired layout

The final `Data Insights` sheet should have this general structure:

1. A title row at the top spanning all visible table columns
2. A left subtitle for `Counter Cards`
3. A left table grouped by class description / occasion-style grouping
4. A right subtitle for `Other Products`
5. A right table starting in column `I`

Both tables should use the same headers and style pattern as the existing occasion tables.

## Recommended implementation approach

### 1. Inspect the current `Data Insights` builder

Find the code that currently builds the `Data Insights` sheet and identify:

- how the title range is defined
- how the left-side tables are grouped and written
- how column widths and row placement are calculated
- where the existing occasion-based layout is assembled

### 2. Split the source data into two groups

Add a clear separation between:

- counter cards
- all other products

The grouping should happen before sheet writing so the table-building logic can stay simple.

### 3. Reuse the existing table-writing pattern

Use the same layout and styling rules as the occasion tables for both sides:

- same headers
- same row style treatment
- same numeric/date formatting behavior
- same grouping logic where applicable

If the existing code already builds one table at a time, factor that into a reusable helper and call it for both left and right sections.

### 4. Extend the title range

Adjust the merged title cell range so it spans both table regions.

This should preserve the current title style while widening it to match the new two-table layout.

### 5. Add section subtitles

Write the section labels directly into the sheet:

- `Counter Cards` on the left
- `Other Products` on the right

Ensure they align visually with the corresponding table regions.

### 6. Place the right-side table in column I

Start the new table in column `I` so the two sections are clearly separated.

If necessary, update spacing, widths, and merged ranges so the sheet remains readable and balanced.

### 7. Verify output formatting

Check that:

- the title spans the full width
- both subtitles appear in the correct places
- the new right-side table begins at column `I`
- the table styling matches the existing occasion tables
- the sheet still renders cleanly with the wider layout

## Implementation notes

- Prefer small helper functions if the current `Data Insights` writer is dense.
- Avoid changing unrelated workbook behavior.
- Keep existing occasion-table formatting intact unless the new layout requires small alignment adjustments.
- If a shared table builder already exists, extend it instead of duplicating logic.

## Checklist

### Planning

- [ ] Locate the current `Data Insights` sheet writer
- [ ] Identify how the existing occasion tables are built
- [ ] Determine how the title range is currently calculated
- [ ] Decide how to split `Counter Cards` from `Other Products`

### Layout changes

- [ ] Extend the title range to cover the new right-side table
- [ ] Add a `Counter Cards` subtitle on the left
- [ ] Add an `Other Products` subtitle on the right
- [ ] Start the new table in column `I`
- [ ] Preserve the existing left-side occasion tables

### Table building

- [ ] Reuse the existing table column layout
- [ ] Keep the same row formatting and grouping style
- [ ] Ensure both tables use consistent styling
- [ ] Avoid duplicating logic if a shared helper can be introduced

### Verification

- [ ] Confirm the `Data Insights` title spans both table regions
- [ ] Confirm the right table starts in column `I`
- [ ] Confirm the section subtitles render correctly
- [ ] Confirm the new table includes all non-counter-card products
- [ ] Confirm existing counter-card/occasion output still works

## Status

Implemented. The `Data Insights` layout update has been completed and validated with `go test ./...`.
