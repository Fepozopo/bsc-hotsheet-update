package hotsheet

import (
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

const mtoSheetName = "MTO"

var mtoHeaders = [...]string{
	"SKU", "Available Quantity", "Forecast Demand", "Projected Stockout Month",
	"MTO", "History Coverage", "Class Description", "Description",
	"Occasion", "Foil", "Card Size",
}

// writeMTOSheet writes active items' 12-month BSC demand and 24-month stockout
// estimates into f as of the history report's run date. Class descriptions use
// the standard SKU prefix rules and numeric MTO cells use MTO YTD's light colors.
// The returned error identifies the first worksheet operation that fails.
func writeMTOSheet(f *excelize.File, entries []*inventoryEntry, asOf time.Time) error {
	if asOf.IsZero() {
		return fmt.Errorf("cannot create MTO sheet without a sales history run date")
	}
	if _, err := f.NewSheet(mtoSheetName); err != nil {
		return fmt.Errorf("failed to create MTO sheet: %w", err)
	}
	headerStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(), Border: thinBlackBorder(),
		Fill: patternFill(standardHeaderFill), Font: boldFont(),
	})
	if err != nil {
		return fmt.Errorf("failed to create MTO header style: %w", err)
	}
	widths := [...]float64{20, 21, 20, 28, 13, 25, 22, 35, 20, 13, 13}
	for col, header := range mtoHeaders {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		if err := f.SetCellValue(mtoSheetName, cell, header); err != nil {
			return fmt.Errorf("failed to write MTO header %s: %w", cell, err)
		}
		if err := f.SetCellStyle(mtoSheetName, cell, cell, headerStyle); err != nil {
			return fmt.Errorf("failed to style MTO header %s: %w", cell, err)
		}
		name, _ := excelize.ColumnNumberToName(col + 1)
		if err := f.SetColWidth(mtoSheetName, name, name, widths[col]); err != nil {
			return fmt.Errorf("failed to set MTO column %s width: %w", name, err)
		}
	}
	// Header comments keep the source date and assumptions visible without
	// displacing the filterable first-row column headers.
	comments := [...]excelize.Comment{
		{Cell: "C1", Author: "Hotsheet Generator", Text: "Projected BSC units sold for 12 calendar months after the report run date " + asOf.Format("Jan 2, 2006") + ". Recent completed same-month sales are weighted toward newer years. Partial months are estimated uniformly by day."},
		{Cell: "E1", Author: "Hotsheet Generator", Text: fmt.Sprintf("Available = on hand + all undated POs - sales orders - backorders. MTO simulates BSC units sold for up to %d months. Undated POs are assumed available immediately; stockout within a month assumes uniform demand.", mtoHorizonMonths)},
		{Cell: "F1", Author: "Hotsheet Generator", Text: "Completed BSC sales months and years from the first observed positive-sale month through the last completed month. Zero-sale months after that first sale count. Missing years and partial months do not."},
	}
	for _, comment := range comments {
		if err := f.AddComment(mtoSheetName, comment); err != nil {
			return fmt.Errorf("failed to annotate MTO header %s: %w", comment.Cell, err)
		}
	}

	bodyStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(), Border: thinBlackBorder(),
		Fill: patternFill(standardSheetCellFillColor("", 0, -1, -1, 0, 0, nil)),
	})
	if err != nil {
		return fmt.Errorf("failed to create MTO cell style: %w", err)
	}
	// Forecast units and MTO can be fractional, even when source sales are whole
	// units, because monthly history is averaged and partial months are prorated.
	numberFormat := "#,##0.0"
	numberStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(), Border: thinBlackBorder(),
		Fill:         patternFill(standardSheetCellFillColor("", 0, -1, -1, 0, 0, nil)),
		CustomNumFmt: &numberFormat,
	})
	if err != nil {
		return fmt.Errorf("failed to create MTO forecast style: %w", err)
	}
	// There are only three numeric YTD bands; reuse each Excel style across rows.
	mtoStyles := make(map[string]int, 3)
	const mtoColumnIdx = 4 // Column E, zero-based like standardSheetCellFillColor.

	for index, row := range buildMTORows(entries, asOf) {
		rowNum := index + 2
		classDesc := applyStandardDisplayClassPrefix(row.item)
		mtoValue := interface{}(row.mto)
		switch row.state {
		case mtoBeyondHorizon:
			mtoValue = fmt.Sprintf(">%d", mtoHorizonMonths)
		case mtoInsufficientHistory:
			mtoValue = "Insufficient history"
		}
		values := [...]interface{}{
			row.item.SKU, row.available, "", row.stockout, mtoValue, row.coverage,
			classDesc, row.item.Description, row.item.Occasion, row.item.Foil, row.item.CardSize,
		}
		if row.hasDemand {
			values[2] = row.demand12
		}
		for col, value := range values {
			cell, _ := excelize.CoordinatesToCellName(col+1, rowNum)
			if err := f.SetCellValue(mtoSheetName, cell, value); err != nil {
				return fmt.Errorf("failed to write MTO cell %s: %w", cell, err)
			}
		}
		if err := f.SetCellStyle(mtoSheetName, fmt.Sprintf("A%d", rowNum), fmt.Sprintf("K%d", rowNum), bodyStyle); err != nil {
			return fmt.Errorf("failed to style MTO row %d: %w", rowNum, err)
		}
		if row.hasDemand {
			if err := f.SetCellStyle(mtoSheetName, fmt.Sprintf("C%d", rowNum), fmt.Sprintf("C%d", rowNum), numberStyle); err != nil {
				return fmt.Errorf("failed to format MTO forecast row %d: %w", rowNum, err)
			}
		}
		if row.state == mtoStockout {
			fill := standardSheetCellFillColor(row.item.Status, mtoColumnIdx, mtoColumnIdx, -1, row.mto, 0, row.mto)
			style, ok := mtoStyles[fill]
			if !ok {
				style, err = f.NewStyle(&excelize.Style{
					Alignment: centeredAlignment(), Border: thinBlackBorder(),
					Fill: patternFill(fill), CustomNumFmt: &numberFormat,
				})
				if err != nil {
					return fmt.Errorf("failed to create MTO YTD color style: %w", err)
				}
				mtoStyles[fill] = style
			}
			cell := fmt.Sprintf("E%d", rowNum)
			if err := f.SetCellStyle(mtoSheetName, cell, cell, style); err != nil {
				return fmt.Errorf("failed to format MTO value row %d: %w", rowNum, err)
			}
		}
	}
	// Excelize adds an Excel dropdown to every column header for filtering.
	if err := f.AutoFilter(mtoSheetName, "A1:K1", nil); err != nil {
		return fmt.Errorf("failed to set MTO autofilter: %w", err)
	}
	return nil
}
