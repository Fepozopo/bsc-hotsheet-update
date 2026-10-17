package hotsheet

import (
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

const dataInsightsSheetName = "Data Insights"

// dataInsightsStyles holds Excelize style IDs shared by both outlook areas.
type dataInsightsStyles struct {
	title, section, header, data, currency, total, currencyTotal int
}

// writeDataInsightsSheet creates a grouped sales outlook: Counter Cards by
// occasion on the left, other products by class and occasion on the right.
// hasHistory inserts PY Sales to Date before full-year PY Sales and enables
// completed-month forecasting using the sales report's runDate, never the
// machine clock. Each table total is forecast from its combined monthly history.
// Excelize writes scalar values and formatting; percentages are explicit text.
func writeDataInsightsSheet(f *excelize.File, entries []*inventoryEntry, hasHistory bool, runDate time.Time) error {
	cards, other := buildDataInsightsTables(entries, hasHistory, runDate)
	if _, err := f.NewSheet(dataInsightsSheetName); err != nil {
		return fmt.Errorf("create Data Insights sheet: %w", err)
	}
	styles, err := newDataInsightsStyles(f)
	if err != nil {
		return err
	}

	headers := []string{"Occasion", "Holiday / Date", "YTD Sales"}
	widths := []float64{26, 22, 16}
	if hasHistory {
		headers = append(headers, "PY Sales to Date")
		widths = append(widths, 19)
	}
	headers = append(headers, "PY Sales", "Projected YoY")
	widths = append(widths, 16, 39)
	// Leave a blank separator column even when the optional history column
	// expands both tables; no fixed column letters can overlap the two areas.
	rightStart := len(headers) + 2
	lastColumn := rightStart + len(headers) - 1
	for _, start := range []int{1, rightStart} {
		for index, width := range widths {
			col, err := excelize.ColumnNumberToName(start + index)
			if err != nil {
				return err
			}
			if err := f.SetColWidth(dataInsightsSheetName, col, col, width); err != nil {
				return err
			}
		}
	}
	if err := writeDataInsightsHeading(f, 1, lastColumn, 1, "Data Insights", styles.title); err != nil {
		return err
	}
	note := "No monthly sales history supplied; projections require at least two completed months."
	if hasHistory {
		note = fmt.Sprintf("Sales history run: %s. Current-month sales are excluded from projections; YTD Sales and PY Sales retain inventory-report values.", runDate.Format("Jan 2, 2006"))
	}
	if err := writeDataInsightsHeading(f, 1, lastColumn, 2, note, styles.data); err != nil {
		return err
	}
	if err := writeDataInsightsHeading(f, rightStart, lastColumn, 3, "Other Products", styles.section); err != nil {
		return err
	}
	if _, err := writeDataInsightsTable(f, 1, 4, "Counter Cards", headers, cards, hasHistory, styles); err != nil {
		return err
	}
	row := 4
	for _, class := range sortedOtherProductsDataInsightsClasses(other) {
		next, err := writeDataInsightsTable(f, rightStart, row, class, headers, other[class], hasHistory, styles)
		if err != nil {
			return err
		}
		row = next + 1
	}
	return nil
}

// newDataInsightsStyles registers the shared text, currency, and heading formats
// with Excelize, returning their IDs or the first style-registration error.
func newDataInsightsStyles(f *excelize.File) (dataInsightsStyles, error) {
	var ids dataInsightsStyles
	definitions := []struct {
		id    *int
		style excelize.Style
	}{
		{&ids.title, excelize.Style{Font: boldFont(14), Alignment: centeredAlignment()}},
		{&ids.section, excelize.Style{Font: boldFont(), Alignment: centeredAlignment(), Border: thinBlackBorder(), Fill: patternFill(dataInsightsSectionFill)}},
		{&ids.header, excelize.Style{Font: boldFont(), Alignment: centeredAlignment(), Border: thinBlackBorder(), Fill: patternFill(standardHeaderFill)}},
		{&ids.data, excelize.Style{Alignment: centeredAlignment(), Border: thinBlackBorder()}},
		{&ids.currency, excelize.Style{Alignment: centeredAlignment(), Border: thinBlackBorder(), CustomNumFmt: currencyNumFmt()}},
		{&ids.total, excelize.Style{Font: boldFont(), Alignment: centeredAlignment(), Border: thinBlackBorder(), Fill: patternFill(dataInsightsTotalFill)}},
		{&ids.currencyTotal, excelize.Style{Font: boldFont(), Alignment: centeredAlignment(), Border: thinBlackBorder(), Fill: patternFill(dataInsightsTotalFill), CustomNumFmt: currencyNumFmt()}},
	}
	for _, definition := range definitions {
		id, err := f.NewStyle(&definition.style)
		if err != nil {
			return ids, fmt.Errorf("create Data Insights style: %w", err)
		}
		*definition.id = id
	}
	return ids, nil
}

// writeDataInsightsHeading writes text into a styled, horizontally merged span
// from startCol to endCol on row; columns are one-based Excel coordinates.
func writeDataInsightsHeading(f *excelize.File, startCol, endCol, row int, text string, style int) error {
	start, err := excelize.CoordinatesToCellName(startCol, row)
	if err != nil {
		return err
	}
	end, err := excelize.CoordinatesToCellName(endCol, row)
	if err != nil {
		return err
	}
	if err := f.SetCellValue(dataInsightsSheetName, start, text); err != nil {
		return err
	}
	if err := f.MergeCell(dataInsightsSheetName, start, end); err != nil {
		return err
	}
	return f.SetCellStyle(dataInsightsSheetName, start, end, style)
}

// writeDataInsightsTable writes table's title, headers, occasion rows, and
// independently forecast total at startCol/startRow, returning the first unused
// row or a workbook-write error. hasHistory controls the optional aligned PY
// column for every row and total.
func writeDataInsightsTable(f *excelize.File, startCol, startRow int, title string, headers []string, table dataInsightsTable, hasHistory bool, styles dataInsightsStyles) (int, error) {
	endCol := startCol + len(headers) - 1
	if err := writeDataInsightsHeading(f, startCol, endCol, startRow, title, styles.section); err != nil {
		return 0, err
	}
	values := make([]interface{}, len(headers))
	for index, header := range headers {
		values[index] = header
	}
	if err := writeDataInsightsValues(f, startCol, startRow+1, values, styles.header, 0); err != nil {
		return 0, err
	}
	rowNum := startRow + 2
	for _, row := range table.Rows {
		if err := writeDataInsightsValues(f, startCol, rowNum, dataInsightsValues(row, hasHistory), styles.data, styles.currency); err != nil {
			return 0, err
		}
		rowNum++
	}
	if err := writeDataInsightsValues(f, startCol, rowNum, dataInsightsValues(table.Total, hasHistory), styles.total, styles.currencyTotal); err != nil {
		return 0, err
	}
	return rowNum + 1, nil
}

// dataInsightsValues returns row's scalar worksheet values in header order,
// inserting PY Sales to Date before full-year PY Sales when hasHistory is true.
func dataInsightsValues(row dataInsightsRow, hasHistory bool) []interface{} {
	values := []interface{}{row.Occasion, row.Date, row.DollarSoldYTD}
	if hasHistory {
		values = append(values, row.DollarSoldPYToDate)
	}
	return append(values, row.DollarSoldPY, dataInsightsYoYDisplay(row))
}

// writeDataInsightsValues writes a contiguous scalar row using Excelize's
// SetSheetRow and applies its style. A nonzero currencyStyle formats only sales
// columns, leaving the mixed percentage/reason text in the last column intact.
func writeDataInsightsValues(f *excelize.File, col, row int, values []interface{}, style, currencyStyle int) error {
	start, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	end, err := excelize.CoordinatesToCellName(col+len(values)-1, row)
	if err != nil {
		return err
	}
	if err := f.SetSheetRow(dataInsightsSheetName, start, &values); err != nil {
		return fmt.Errorf("write Data Insights row %s: %w", start, err)
	}
	if err := f.SetCellStyle(dataInsightsSheetName, start, end, style); err != nil {
		return err
	}
	if currencyStyle != 0 {
		currencyStart, err := excelize.CoordinatesToCellName(col+2, row)
		if err != nil {
			return err
		}
		currencyEnd, err := excelize.CoordinatesToCellName(col+len(values)-2, row)
		if err != nil {
			return err
		}
		return f.SetCellStyle(dataInsightsSheetName, currencyStart, currencyEnd, currencyStyle)
	}
	return nil
}
