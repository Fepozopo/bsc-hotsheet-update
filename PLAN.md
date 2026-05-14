# PLAN

## Goal

Add a new workbook sheet called `Data Insights` that summarizes sales by occasion for:

- Spring occasions
- Winter occasions
- Everyday products

The sheet should be designed as a reusable reporting output so we can later add more insights without rewriting the workbook generation flow.

## What this sheet should do

- Create one row per occasion
- Split the rows into three tables:
  1. Spring
  2. Winter
  3. Everyday products
- Only include products whose category matches `Counter Cards` using exact text matching
- Use `NO OCCASION` when an occasion is missing
- Add a total row at the bottom of each table

### Final columns

For all three tables, the left-to-right column order should be:

1. Occasion
2. Date
3. YTD Sales
4. PY Sales
5. Final status column

### Final status column by section

- **Spring**: `Status / YoY`
- **Winter**: `Status / YoY`
- **Everyday products**: `Projected YoY`

### Table behavior by section

- **Spring and Winter**
  - The final column should show a status plus YoY context
  - The status should be either `COMPLETE` or `IN PROGRESS`
  - Sort rows by date
- **Everyday products**
  - The final column should show a projected YoY percentage
  - `NO OCCASION` rows belong here
  - Sort rows alphabetically because these rows do not have dates
  - Use `N/A` in the Date column for these rows

### Date formatting

- Use a full month/day format like `February 14`
- `Graduation` should use month-only formatting
- The only month-only value requested so far is `June`
- Everyday rows should use `N/A` in the Date column

### Totals

- Sum the YTD, PY, and projected columns as needed
- Spring and Winter total rows should leave the status column blank
- Everyday total rows should show a total projected percentage in the final column
- The projected percentage total for Everyday should be calculated from the summed YTD and PY totals

## How the inventory data flows

The current implementation builds an inventory map before it generates each workbook, then `CreateFromReports()` materializes those rows into the populated `vals` data used for reporting:

- `invMap` is a `map[string]*entry`
- The key is the SKU
- Each value is a pointer to an `entry` struct that stores the parsed inventory row
- `vals` is the reporting-ready row set derived from those entries and updated with PO data

### What gets stored in each `entry`

The inventory row parser fills in fields like these, and those fields are then carried forward into `vals`, which is what the reporting logic should read from:

- `SKU`
- `ProductLine`
- `ClassDesc`
- `Status`
- inventory quantities such as on-hand, on PO, on SO, on BO
- year-to-date and previous-year sales values
- `Foil`
- `Occasion`
- `Description`
- `UPC`
- `RoyaltyCode`
- dollar sales fields

### How the map is populated

- The inventory workbook is read first
- Each SKU row is parsed into an `entry`
- That `entry` is stored in `invMap` by SKU
- The PO workbook is then read, and matching SKUs in `invMap` are updated with PO information
- `CreateFromReports()` then uses those enriched entries to build `vals` for downstream reporting

### Why this matters for `Data Insights`

For the new sheet, `invMap` is the underlying inventory store, but `CreateFromReports()` should be treated as the place where those entries are materialized into the fully populated `vals` data used for reporting.

That means the new sheet should use the populated row values from `vals` so each item has the correct:

- class description
- PO values
- sales order values
- total available
- MTO values
- other fields already assembled from the inventory and PO data

The implementation should then:

- start from the populated `vals` rows
- filter to exact `Counter Cards` rows
- group rows by occasion
- derive the section using the existing occasion mapping logic
- build one summary row per occasion

### Occasion grouping logic already in the code

The helper `mapOccasion` currently maps raw occasion text into one of three sections:

- `Spring`
- `Winter`
- `Everyday`

It does this by checking the occasion text against token lists.

For the new sheet, we will still use that same section grouping, but we will hard-code a date map for each Spring and Winter occasion so the output date is controlled explicitly.

## Proposed implementation approach

### Phase 1: Data modeling

- Identify the data needed from the populated `vals` rows to populate the new sheet
- Define a clear structure for a row in the insights tables
- Determine how to group the rows into spring, winter, and everyday sections
- Filter `vals` rows so the sheet only includes products in the exact `Counter Cards` category

### Phase 2: Sheet generation helper

- Add a dedicated function for building the `Data Insights` sheet
- Keep the function focused on workbook output only
- Reuse existing style helpers if they already exist

### Phase 3: Table layout

- Create a title row for the sheet
- Render the spring table
- Render the winter table
- Render the everyday products table
- Add a total row to each table

### Phase 4: Verification

- Confirm the sheet is added to the workbook
- Confirm rows appear in the intended order
- Confirm totals match the populated `vals` rows
- Confirm no symbols or emoji are present in the output

## Suggested function responsibilities

### New helper function

A new function should probably be responsible for:

- Creating the sheet
- Writing section headers
- Writing rows for each table
- Writing total rows
- Applying any formatting needed for headers and totals

### Existing orchestration function

The current workbook creation flow should probably remain responsible for:

- Gathering report data and building `vals`
- Calling the new `Data Insights` sheet builder with the populated rows
- Preserving current workbook generation behavior

## Potential data structure

We may want a reusable row model that contains:

- Occasion or category label
- Date
- YTD sales
- PY sales
- Projected sales
- Status or YoY text
- Section name

This would make it easier to reuse the inventory or report breakdown elsewhere later.

## Acceptance criteria

- A new sheet named `Data Insights` is created in the workbook
- The sheet contains three distinct tables
- Each table has a total row
- Spring, winter, and everyday items are all represented
- The date column replaces the peak season concept
- No symbols or emoji are used in the sheet
- The implementation only includes products in the exact `Counter Cards` category
- The implementation is modular enough to reuse the underlying data later

## Open implementation notes

- The Spring and Winter `Status / YoY` examples should follow the pattern `COMPLETE: -77% YoY` or `IN PROGRESS: +2% so far`
- Everyday rows should use `N/A` in the Date column
- The total projected percentage for Everyday rows should be calculated directly from summed YTD and PY totals
- The Spring and Winter occasion dates should come from a hard-coded date map
- We should confirm whether the sheet layout should mirror the screenshot closely, or only follow the same general structure

## Concrete build checklist

### Data preparation

- [x] Confirm the populated `vals` rows from `CreateFromReports()` are the source of truth for occasion-level rows, with `invMap` as the input data source
- [x] Filter `vals` rows to exact `Counter Cards` category matches only
- [x] Normalize missing occasions to `NO OCCASION`
- [x] Group `vals` rows so there is exactly one output row per occasion
- [x] Split grouped rows into `Spring`, `Winter`, and `Everyday` buckets using the existing occasion mapping logic

### Occasion and date handling

- [x] Add a hard-coded date map for each Spring and Winter occasion
- [x] Use full month/day values for Spring and Winter dates, except `Graduation`
- [x] Use `June` for `Graduation`
- [x] Use `N/A` in the Date column for all Everyday rows
- [x] Confirm Everyday rows with `NO OCCASION` remain in the Everyday table

### Sorting

- [x] Sort Spring rows by date
- [x] Sort Winter rows by date
- [x] Sort Everyday rows alphabetically by occasion
- [x] Make sure sorting is done after grouping and before writing rows

### Column layout

- [x] Keep the final sheet column order as: Occasion, Date, YTD Sales, PY Sales, final status column
- [x] Use `Status / YoY` as the final column name for Spring and Winter tables
- [x] Use `Projected YoY` as the final column name for the Everyday table
- [x] Ensure the final column contains percentages where required

### Row calculations

- [x] Calculate one row per occasion from the populated `vals` rows
- [x] Sum YTD sales per occasion
- [x] Sum PY sales per occasion
- [x] Sum projected values per occasion where applicable
- [x] For Spring and Winter rows, format the final column as `COMPLETE: -77% YoY` or `IN PROGRESS: +2% so far`
- [x] For Everyday rows, format the final column as a projected YoY percentage

### Total row logic

- [x] Add a total row at the bottom of each table
- [x] Leave the Spring and Winter status cells blank in total rows
- [x] Calculate the Everyday total projected percentage from summed YTD and PY totals
- [x] Put the Everyday total projected percentage in the final column of the total row

### Workbook creation

- [x] Add a new helper function to create the `Data Insights` sheet
- [x] Insert the sheet at the end of the workbook
- [x] Write the sheet title and section headers
- [x] Render the Spring, Winter, and Everyday tables in order
- [x] Apply the existing workbook styling conventions where practical
- [x] Avoid symbols and emoji in the new sheet

### Verification

- [x] Confirm the sheet is added as `Data Insights`
- [x] Confirm there is one row per occasion
- [x] Confirm `Counter Cards` is matched with exact text
- [x] Confirm `Graduation` shows `June`
- [x] Confirm `NO OCCASION` appears only in the Everyday table
- [x] Confirm the total rows are present and values are correct
- [x] Confirm the final column text matches the requested formats

## Status

Implemented the new helper function and wired it into the workbook generation flow.
