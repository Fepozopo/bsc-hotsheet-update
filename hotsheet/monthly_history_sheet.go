package hotsheet

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

const monthlyHistorySheetName = "Monthly History"

var monthlyHistoryHeaders = [17]string{
	"Item Code", "Description", "Year", "Metric",
	"Jan", "Feb", "Mar", "Apr", "May", "Jun",
	"Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Status",
}

// writeMonthlyHistorySheet adds entries' BSC sales and issued records to f as a
// product-line tab, reusing a reserved empty tab if one already exists. Rows are
// streamed in source order to avoid retaining a second in-memory copy of the
// history. Each row is one SKU/year/metric with status last for filtering.
// Call this after workbook operations that read all sheets, such as hiding helper
// tabs, and save f before reading this sheet. The function returns an error if
// the sheet cannot be created, written, or flushed.
func writeMonthlyHistorySheet(f *excelize.File, entries []*inventoryEntry) error {
	if _, err := f.NewSheet(monthlyHistorySheetName); err != nil {
		return fmt.Errorf("failed to create monthly history sheet: %w", err)
	}

	// Configure the filter before streaming so Excelize includes this worksheet
	// metadata on Flush without mixing normal cell writes with streamed rows.
	if err := f.AutoFilter(monthlyHistorySheetName, "A1:Q1", nil); err != nil {
		return fmt.Errorf("failed to set monthly history autofilter: %w", err)
	}
	// Excelize serializes rows as they arrive and spills large buffers to disk.
	// Column widths must be set before the first SetRow call.
	stream, err := f.NewStreamWriter(monthlyHistorySheetName)
	if err != nil {
		return fmt.Errorf("failed to stream monthly history sheet: %w", err)
	}
	for col, header := range monthlyHistoryHeaders {
		width := 12.0
		switch header {
		case "Item Code":
			width = 20
		case "Description":
			width = 35
		case "Metric":
			width = 24
		case "Status":
			width = 15
		}
		if err := stream.SetColWidth(col+1, col+1, width); err != nil {
			return fmt.Errorf("failed to set monthly history column %d width: %w", col+1, err)
		}
	}

	headerStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(),
		Border:    thinBlackBorder(),
		Fill:      patternFill(standardHeaderFill),
		Font:      boldFont(),
	})
	if err != nil {
		return fmt.Errorf("failed to create monthly history header style: %w", err)
	}
	var headers [17]interface{}
	for col, header := range monthlyHistoryHeaders {
		headers[col] = header
	}
	if err := stream.SetRow("A1", headers[:], excelize.RowOpts{StyleID: headerStyle}); err != nil {
		return fmt.Errorf("failed to write monthly history headers: %w", err)
	}

	// Reuse the standard sheet's status colors; create styles once rather than once
	// per cell so multi-year workbooks do not repeat expensive style registration.
	var rowStyles, currencyStyles, percentStyles [3]int
	for i, status := range []string{"Carryover", "Rundown", "Discontinued"} {
		fill := standardSheetCellFillColor(status, 0, -1, -1, 0, 0, nil)
		for kind := range 3 {
			styleDef := &excelize.Style{
				Alignment: centeredAlignment(),
				Border:    thinBlackBorder(),
				Fill:      patternFill(fill),
			}
			switch kind {
			case 1:
				styleDef.CustomNumFmt = currencyNumFmt()
			case 2:
				// The source stores 85 for 85%, not 0.85 as Excel's percent format expects.
				format := `0.00"%"`
				styleDef.CustomNumFmt = &format
			}
			style, err := f.NewStyle(styleDef)
			if err != nil {
				return fmt.Errorf("failed to create monthly history cell style: %w", err)
			}
			switch kind {
			case 0:
				rowStyles[i] = style
			case 1:
				currencyStyles[i] = style
			case 2:
				percentStyles[i] = style
			}
		}
	}

	rowNum := 2
	for _, item := range entries {
		for _, record := range item.SalesRecords {
			statusIndex := 0
			switch item.Status {
			case "Rundown":
				statusIndex = 1
			case "Discontinued":
				statusIndex = 2
			}
			style := rowStyles[statusIndex]
			switch record.Metric {
			case "Dollars Sold", "Cost of Goods Sold":
				style = currencyStyles[statusIndex]
			case "Gross Profit Percent":
				style = percentStyles[statusIndex]
			}
			values := [17]interface{}{item.SKU, item.Description, record.Year, record.Metric}
			// RowOpts applies status shading to the entire row; Cell overrides
			// preserve currency/percentage display only in the monthly columns.
			for period, value := range record.Periods {
				values[period+4] = excelize.Cell{StyleID: style, Value: value}
			}
			values[16] = item.Status
			if err := stream.SetRow(fmt.Sprintf("A%d", rowNum), values[:], excelize.RowOpts{StyleID: rowStyles[statusIndex]}); err != nil {
				return fmt.Errorf("failed to write monthly history row %d: %w", rowNum, err)
			}
			rowNum++
		}
	}

	// Flush completes the sheet XML and reports any temporary-file write error.
	if err := stream.Flush(); err != nil {
		return fmt.Errorf("failed to flush monthly history sheet: %w", err)
	}
	return nil
}
