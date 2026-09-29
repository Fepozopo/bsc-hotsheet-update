package hotsheet

import (
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	mtoSheetName         = "MTO"
	mtoScenarioSheetName = "MTO Forecast Data"
	maxProposedPOUnits   = 2147483647
)

var mtoHeaders = [...]string{
	"SKU", "Available Quantity", "Forecast Demand", "Projected Stockout Month",
	"MTO", "Proposed PO Units", "MTO with Proposed PO", "Stockout Month with Proposed PO",
	"History Coverage", "Class Description", "Occasion", "Foil", "Description", "Card Size",
}

// writeMTOSheet writes active items' 12-month BSC demand and 24-month stockout
// estimates into f as of the history report's run date. Missing calendar months
// use the observed-month mean once two completed months are available. Hidden
// forecast segments power live PO what-if formulas without altering baseline
// quantities. Class descriptions use standard SKU prefixes and numeric baseline
// MTO cells use MTO YTD's light colors.
// The returned error identifies the first worksheet operation that fails.
func writeMTOSheet(f *excelize.File, entries []*inventoryEntry, asOf time.Time) error {
	if asOf.IsZero() {
		return fmt.Errorf("cannot create MTO sheet without a sales history run date")
	}
	if _, err := f.NewSheet(mtoSheetName); err != nil {
		return fmt.Errorf("failed to create MTO sheet: %w", err)
	}
	if _, err := f.NewSheet(mtoScenarioSheetName); err != nil {
		return fmt.Errorf("failed to create MTO forecast data sheet: %w", err)
	}
	headerStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(), Border: thinBlackBorder(),
		Fill: patternFill(standardHeaderFill), Font: boldFont(),
	})
	if err != nil {
		return fmt.Errorf("failed to create MTO header style: %w", err)
	}
	widths := [...]float64{20, 21, 20, 28, 13, 20, 22, 34, 25, 22, 20, 13, 35, 13}
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
		{Cell: "C1", Author: "Hotsheet Generator", Text: "Projected BSC units sold for 12 calendar months after the report run date " + asOf.Format("Jan 2, 2006") + ". Recent completed same-month sales are weighted toward newer years. With at least two completed sales months, missing calendar months use the average of observed months. Partial months are estimated uniformly by day."},
		{Cell: "E1", Author: "Hotsheet Generator", Text: fmt.Sprintf("Available = on hand + all undated POs - sales orders - backorders. MTO simulates BSC units sold for up to %d months. Undated POs are assumed available immediately; stockout within a month assumes uniform demand.", mtoHorizonMonths)},
		{Cell: "F1", Author: "Hotsheet Generator", Text: "Enter additional, nonnegative whole units to simulate a PO. Existing POs are already included in Available Quantity. Proposed units are assumed available immediately; this input does not create an order."},
		{Cell: "G1", Author: "Hotsheet Generator", Text: "Recalculates MTO with Available Quantity plus Proposed PO Units using the same monthly BSC forecast and 24-month horizon as the baseline MTO. Blank means zero additional units."},
		{Cell: "I1", Author: "Hotsheet Generator", Text: "Completed BSC sales months and years from the first observed positive-sale month through the last completed month. Zero-sale months after that first sale count. Missing years and partial months do not. Forecasts based on fewer than 12 observed calendar months use the completed-month average for missing calendar months."},
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
	inputStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(), Border: thinBlackBorder(),
		Fill: patternFill("FFF2CC"),
	})
	if err != nil {
		return fmt.Errorf("failed to create editable PO input style: %w", err)
	}
	// There are only three numeric YTD bands; reuse each Excel style across rows.
	mtoStyles := make(map[string]int, 3)
	const mtoColumnIdx = 4 // Column E, zero-based like standardSheetCellFillColor.

	dataRow := 2
	start := asOf.AddDate(0, 0, 1)
	end := monthsAfter(start, mtoHorizonMonths)
	rows := buildMTORows(entries, asOf)
	for index, row := range rows {
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
			row.item.SKU, row.available, "", row.stockout, mtoValue, "", "", "",
			row.coverage, classDesc, row.item.Occasion, row.item.Foil, row.item.Description, row.item.CardSize,
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
		if err := f.SetCellStyle(mtoSheetName, fmt.Sprintf("A%d", rowNum), fmt.Sprintf("N%d", rowNum), bodyStyle); err != nil {
			return fmt.Errorf("failed to style MTO row %d: %w", rowNum, err)
		}
		if row.hasDemand {
			if err := f.SetCellStyle(mtoSheetName, fmt.Sprintf("C%d", rowNum), fmt.Sprintf("C%d", rowNum), numberStyle); err != nil {
				return fmt.Errorf("failed to format MTO forecast row %d: %w", rowNum, err)
			}
		}
		if err := f.SetCellStyle(mtoSheetName, fmt.Sprintf("F%d", rowNum), fmt.Sprintf("F%d", rowNum), inputStyle); err != nil {
			return fmt.Errorf("failed to style proposed PO input row %d: %w", rowNum, err)
		}
		if err := f.SetCellStyle(mtoSheetName, fmt.Sprintf("G%d", rowNum), fmt.Sprintf("G%d", rowNum), numberStyle); err != nil {
			return fmt.Errorf("failed to format proposed PO MTO row %d: %w", rowNum, err)
		}
		if err := writeMTOScenarioFormulas(f, row, rowNum, &dataRow, start, end, asOf); err != nil {
			return err
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
	if len(rows) > 0 {
		validation := excelize.NewDataValidation(true)
		validation.Sqref = fmt.Sprintf("F2:F%d", len(rows)+1)
		if err := validation.SetRange(0, maxProposedPOUnits, excelize.DataValidationTypeWhole, excelize.DataValidationOperatorBetween); err != nil {
			return fmt.Errorf("failed to configure proposed PO validation: %w", err)
		}
		validation.SetInput("Proposed PO units", "Enter additional units; blank means zero. Assumed available immediately.")
		validation.SetError(excelize.DataValidationErrorStyleStop, "Invalid PO quantity", "Enter a nonnegative whole number.")
		if err := f.AddDataValidation(mtoSheetName, validation); err != nil {
			return fmt.Errorf("failed to validate proposed PO units: %w", err)
		}
		// Excel conditional formatting updates the scenario MTO colors when the
		// input changes; plain cell styles would retain the initial color.
		var rules []excelize.ConditionalFormatOptions
		for _, band := range []struct {
			criteria string
			value    float64
		}{
			{"AND(ISNUMBER(G2),G2<=1)", 0},
			{"AND(ISNUMBER(G2),G2>1,G2<=3)", 2},
			{"AND(ISNUMBER(G2),G2>3)", 4},
		} {
			fill := standardSheetCellFillColor("", 4, 4, -1, band.value, 0, band.value)
			style, err := f.NewConditionalStyle(&excelize.Style{Fill: patternFill(fill)})
			if err != nil {
				return fmt.Errorf("failed to create proposed PO MTO color: %w", err)
			}
			rules = append(rules, excelize.ConditionalFormatOptions{Type: "formula", Criteria: band.criteria, Format: &style})
		}
		if err := f.SetConditionalFormat(mtoSheetName, fmt.Sprintf("G2:G%d", len(rows)+1), rules); err != nil {
			return fmt.Errorf("failed to color proposed PO MTO: %w", err)
		}
	}
	// Excelize adds an Excel dropdown to every column header for filtering.
	if err := f.AutoFilter(mtoSheetName, "A1:N1", nil); err != nil {
		return fmt.Errorf("failed to set MTO autofilter: %w", err)
	}
	if err := f.SetSheetVisible(mtoScenarioSheetName, false); err != nil {
		return fmt.Errorf("failed to hide MTO forecast data: %w", err)
	}
	// Excelize stores formulas without their calculated results; request a full
	// calculation on open so the blank-input scenario displays the baseline.
	auto, recalculate := "auto", true
	if err := f.SetCalcProps(&excelize.CalcPropsOptions{CalcMode: &auto, FullCalcOnLoad: &recalculate}); err != nil {
		return fmt.Errorf("failed to enable MTO scenario recalculation: %w", err)
	}
	return nil
}

// writeMTOScenarioFormulas records the same prorated monthly segments as the Go
// stockout simulation and links a row's proposed-PO input to its live Excel results.
// dataRow advances past the segments used; missing history remains unforecastable.
func writeMTOScenarioFormulas(f *excelize.File, row mtoForecastRow, rowNum int, dataRow *int, start, end, asOf time.Time) error {
	quantity := fmt.Sprintf("B%d+F%d", rowNum, rowNum)
	result := fmt.Sprintf("IF(%s<=0,0,\"Insufficient history\")", quantity)
	month := fmt.Sprintf("IF(%s<=0,\"%s\",\"Insufficient history\")", quantity, asOf.Format("Jan 2006"))
	if row.hasDemand {
		first := *dataRow
		cumulative, elapsed := 0.0, 0.0
		for monthStart := time.Date(start.Year(), start.Month(), 1, 0, 0, 0, 0, time.UTC); monthStart.Before(end); monthStart = monthStart.AddDate(0, 1, 0) {
			monthEnd := monthStart.AddDate(0, 1, 0)
			segmentStart, segmentEnd := monthStart, monthEnd
			if segmentStart.Before(start) {
				segmentStart = start
			}
			if segmentEnd.After(end) {
				segmentEnd = end
			}
			fraction := segmentEnd.Sub(segmentStart).Hours() / monthEnd.Sub(monthStart).Hours()
			units := row.monthly[int(monthStart.Month())-1]
			values := [...]interface{}{row.item.SKU, units, cumulative, cumulative + units*fraction, elapsed, monthStart.Format("Jan 2006")}
			for col, value := range values {
				cell, _ := excelize.CoordinatesToCellName(col+1, *dataRow)
				if err := f.SetCellValue(mtoScenarioSheetName, cell, value); err != nil {
					return fmt.Errorf("failed to write MTO forecast data %s: %w", cell, err)
				}
			}
			cumulative += units * fraction
			elapsed += fraction
			*dataRow += 1
		}
		last := *dataRow - 1
		ref := func(col string) string {
			return fmt.Sprintf("'%s'!$%s$%d:$%s$%d", mtoScenarioSheetName, col, first, col, last)
		}
		index := fmt.Sprintf("MIN(COUNTIF(%s,\"<\"&(%s))+1,%d)", ref("D"), quantity, last-first+1)
		// The last cumulative value bounds the simulated horizon.
		result = fmt.Sprintf("IF(%s<=0,0,IF(%s>'%s'!D%d,\">%d\",IFERROR(INDEX(%s,%s)+(%s-INDEX(%s,%s))/INDEX(%s,%s),\">%d\")))",
			quantity, quantity, mtoScenarioSheetName, last, mtoHorizonMonths, ref("E"), index, quantity, ref("C"), index, ref("B"), index, mtoHorizonMonths)
		month = fmt.Sprintf("IF(G%d=0,\"%s\",IF(ISNUMBER(G%d),INDEX(%s,%s),\"Not within %d months\"))",
			rowNum, asOf.Format("Jan 2006"), rowNum, ref("F"), index, mtoHorizonMonths)
	}
	valid := fmt.Sprintf("OR(F%d=\"\",IFERROR(AND(ISNUMBER(F%d),F%d>=0,F%d=INT(F%d)),FALSE))", rowNum, rowNum, rowNum, rowNum, rowNum)
	if err := f.SetCellFormula(mtoSheetName, fmt.Sprintf("G%d", rowNum), fmt.Sprintf("IFERROR(IF(%s,%s,\"Invalid PO units\"),\"Invalid PO units\")", valid, result)); err != nil {
		return fmt.Errorf("failed to set proposed PO MTO formula on row %d: %w", rowNum, err)
	}
	if err := f.SetCellFormula(mtoSheetName, fmt.Sprintf("H%d", rowNum), fmt.Sprintf("IFERROR(IF(%s,%s,\"Invalid PO units\"),\"Invalid PO units\")", valid, month)); err != nil {
		return fmt.Errorf("failed to set proposed PO stockout formula on row %d: %w", rowNum, err)
	}
	return nil
}
