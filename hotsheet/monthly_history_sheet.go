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
// product-line tab. Each row is one SKU/year/metric with status last for filtering;
// it returns an error if the sheet cannot be created or written.
func writeMonthlyHistorySheet(f *excelize.File, entries []*inventoryEntry) error {
	if _, err := f.NewSheet(monthlyHistorySheetName); err != nil {
		return fmt.Errorf("failed to create monthly history sheet: %w", err)
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
	for col, header := range monthlyHistoryHeaders {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		if err := f.SetCellValue(monthlyHistorySheetName, cell, header); err != nil {
			return fmt.Errorf("failed to write monthly history header %s: %w", cell, err)
		}
		if err := f.SetCellStyle(monthlyHistorySheetName, cell, cell, headerStyle); err != nil {
			return fmt.Errorf("failed to style monthly history header %s: %w", cell, err)
		}
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
			values := [17]interface{}{item.SKU, item.Description, record.Year, record.Metric}
			for period, value := range record.Periods {
				values[period+4] = value
			}
			values[16] = item.Status
			for col, value := range values {
				cell, _ := excelize.CoordinatesToCellName(col+1, rowNum)
				if err := f.SetCellValue(monthlyHistorySheetName, cell, value); err != nil {
					return fmt.Errorf("failed to write monthly history cell %s: %w", cell, err)
				}
			}

			statusIndex := 0
			switch item.Status {
			case "Rundown":
				statusIndex = 1
			case "Discontinued":
				statusIndex = 2
			}
			start := fmt.Sprintf("A%d", rowNum)
			end := fmt.Sprintf("Q%d", rowNum)
			if err := f.SetCellStyle(monthlyHistorySheetName, start, end, rowStyles[statusIndex]); err != nil {
				return fmt.Errorf("failed to style monthly history row %d: %w", rowNum, err)
			}
			style := 0
			switch record.Metric {
			case "Dollars Sold", "Cost of Goods Sold":
				style = currencyStyles[statusIndex]
			case "Gross Profit Percent":
				style = percentStyles[statusIndex]
			}
			if style != 0 {
				if err := f.SetCellStyle(monthlyHistorySheetName, fmt.Sprintf("E%d", rowNum), fmt.Sprintf("P%d", rowNum), style); err != nil {
					return fmt.Errorf("failed to format monthly history row %d: %w", rowNum, err)
				}
			}
			rowNum++
		}
	}

	for col, header := range monthlyHistoryHeaders {
		name, _ := excelize.ColumnNumberToName(col + 1)
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
		if err := f.SetColWidth(monthlyHistorySheetName, name, name, width); err != nil {
			return fmt.Errorf("failed to set monthly history column %s width: %w", name, err)
		}
	}
	// Match the season sheets: filter dropdowns on every header, including Status.
	if err := f.AutoFilter(monthlyHistorySheetName, "A1:Q1", nil); err != nil {
		return fmt.Errorf("failed to set monthly history autofilter: %w", err)
	}
	return nil
}
