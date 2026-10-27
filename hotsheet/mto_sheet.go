package hotsheet

import (
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	mtoSheetName            = "MTO"
	mtoScenarioSheetName    = "MTO Forecast Data"
	mtoScenarioFirstDataRow = 2 // Preserve the helper sheet's unused first row.
	maxProposedPOUnits      = 2147483647
)

// mtoHeaders places a stockout date immediately after each numeric MTO result.
var mtoHeaders = [...]string{
	"SKU", "On Hand", "On PO", "Committed", "Available",
	"Forecast Demand", "MTO", "Stockout Date", "Proposed PO Units", "MTO with Proposed PO", "Projected Stockout Date",
	"History Coverage (months)", "Class Description", "Season", "Occasion", "Description", "Foil", "Card Size",
}

// writeMTOSheet writes rows' inventory quantities, 12-month BSC shipped demand,
// numeric history coverage, and 24-month MTO estimates into f as of the report
// run date. Sortable Excel stockout dates follow each MTO result; hidden forecast
// segments power live PO what-if formulas without altering baseline quantities.
// Forecasts use completed-month means for missing seasonal months once two
// observations exist. Designated 2021 SKUs use only February 2026 onward.
// Class descriptions use standard SKU prefixes; Season uses the inventory tabs'
// occasion mapping. Description precedes Foil; numeric MTO uses light YTD colors.
// rows must be the ordered output of buildMTORows for asOf. This writes the
// visible formulas and reserves a hidden helper tab; the caller must stream
// the same rows with writeMTOForecastDataSheet after all sheet-visibility changes.
// The returned error identifies the first worksheet operation that fails.
func writeMTOSheet(f *excelize.File, rows []mtoForecastRow, asOf time.Time) error {
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
	widths := [...]float64{20, 15, 15, 15, 15, 20, 13, 20, 20, 22, 25, 25, 22, 15, 20, 35, 13, 13}
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
		{Cell: "F1", Author: "Hotsheet Generator", Text: "Projected BSC units shipped (sold plus nonnegative issued units, when supplied) for 12 calendar months after the report run date " + asOf.Format("Jan 2, 2006") + ". Recent completed same-month shipments are weighted toward newer years. With at least two completed shipment months, missing calendar months use the average of observed months. Partial months are estimated uniformly by day."},
		{Cell: "G1", Author: "Hotsheet Generator", Text: fmt.Sprintf("Available = on hand + all undated POs - committed (sales orders + backorders). MTO simulates BSC units shipped for up to %d months. Undated POs are assumed available immediately; stockout within a month assumes uniform demand.", mtoHorizonMonths)},
		{Cell: "H1", Author: "Hotsheet Generator", Text: "Estimated depletion date using uniform daily demand within each calendar month, not a guaranteed stockout. A midnight depletion belongs to the new day. Already depleted stock shows the report run date. No estimate beyond 24 months or with insufficient history."},
		{Cell: "I1", Author: "Hotsheet Generator", Text: "Enter additional, nonnegative whole units to simulate a PO. Existing POs are already included in Available. Proposed units are assumed available immediately; this input does not create an order."},
		{Cell: "J1", Author: "Hotsheet Generator", Text: "Recalculates MTO with Available plus Proposed PO Units using the same monthly BSC forecast and 24-month horizon as the baseline MTO. Blank means zero additional units."},
		{Cell: "K1", Author: "Hotsheet Generator", Text: "Estimated depletion date with Proposed PO Units, using the same daily demand assumption as Stockout Date. Recalculates when proposed units change; blank input matches the baseline date. Existing and proposed POs are assumed available immediately."},
		{Cell: "L1", Author: "Hotsheet Generator", Text: "Number of completed BSC shipment months from the first observed positive-shipment month through the last completed month. Zero-shipment months after that first shipment count. Missing years and partial months do not. Forecasts based on fewer than 12 observed calendar months use the completed-month average for missing calendar months. Designated product-line 2021 SKUs exclude all history before February 2026, when the key account moved to custom SKUs."},
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
	// Excelize stores time.Time values as Excel date serials; applying the
	// same date format to baseline values and formulas keeps both sortable.
	dateFormat := "mmm d, yyyy"
	dateStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(), Border: thinBlackBorder(),
		Fill:         patternFill(standardSheetCellFillColor("", 0, -1, -1, 0, 0, nil)),
		CustomNumFmt: &dateFormat,
	})
	if err != nil {
		return fmt.Errorf("failed to create MTO date style: %w", err)
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
	const mtoColumnIdx = 6 // Column G, zero-based like standardSheetCellFillColor.

	dataRow := mtoScenarioFirstDataRow
	segmentCount := len(mtoScenarioSegments(asOf))
	for index, row := range rows {
		rowNum := index + 2
		classDesc := applyStandardDisplayClassPrefix(row.item)
		mtoValue, stockoutValue := interface{}(row.mto), interface{}(row.stockout)
		switch row.state {
		case mtoBeyondHorizon:
			mtoValue = fmt.Sprintf(">%d", mtoHorizonMonths)
			stockoutValue = fmt.Sprintf("Not within %d months", mtoHorizonMonths)
		case mtoInsufficientHistory:
			mtoValue, stockoutValue = "Insufficient history", "Insufficient history"
		}
		values := [...]interface{}{
			row.item.SKU, row.item.OnHand, row.item.OnPO, row.item.OnSO + row.item.OnBO, row.available,
			"", mtoValue, stockoutValue, "", "", "",
			row.coverage, classDesc, mapOccasion(row.item.Occasion), row.item.Occasion, row.item.Description, row.item.Foil, row.item.CardSize,
		}
		if row.hasDemand {
			values[5] = row.demand12
		}
		for col, value := range values {
			cell, _ := excelize.CoordinatesToCellName(col+1, rowNum)
			if err := f.SetCellValue(mtoSheetName, cell, value); err != nil {
				return fmt.Errorf("failed to write MTO cell %s: %w", cell, err)
			}
		}
		if err := f.SetCellStyle(mtoSheetName, fmt.Sprintf("A%d", rowNum), fmt.Sprintf("R%d", rowNum), bodyStyle); err != nil {
			return fmt.Errorf("failed to style MTO row %d: %w", rowNum, err)
		}
		if row.hasDemand {
			if err := f.SetCellStyle(mtoSheetName, fmt.Sprintf("F%d", rowNum), fmt.Sprintf("F%d", rowNum), numberStyle); err != nil {
				return fmt.Errorf("failed to format MTO forecast row %d: %w", rowNum, err)
			}
		}
		if err := f.SetCellStyle(mtoSheetName, fmt.Sprintf("I%d", rowNum), fmt.Sprintf("I%d", rowNum), inputStyle); err != nil {
			return fmt.Errorf("failed to style proposed PO input row %d: %w", rowNum, err)
		}
		if err := f.SetCellStyle(mtoSheetName, fmt.Sprintf("J%d", rowNum), fmt.Sprintf("J%d", rowNum), numberStyle); err != nil {
			return fmt.Errorf("failed to format proposed PO MTO row %d: %w", rowNum, err)
		}
		for _, column := range []string{"H", "K"} {
			cell := fmt.Sprintf("%s%d", column, rowNum)
			if err := f.SetCellStyle(mtoSheetName, cell, cell, dateStyle); err != nil {
				return fmt.Errorf("failed to format MTO date %s: %w", cell, err)
			}
		}
		// Reserve the same contiguous range the late helper-sheet writer uses.
		// Rows without usable demand must not consume helper rows.
		if err := writeMTOScenarioFormula(f, row, asOf, rowNum, dataRow, dataRow+segmentCount-1); err != nil {
			return err
		}
		if row.hasDemand {
			dataRow += segmentCount
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
			cell := fmt.Sprintf("G%d", rowNum)
			if err := f.SetCellStyle(mtoSheetName, cell, cell, style); err != nil {
				return fmt.Errorf("failed to format MTO value row %d: %w", rowNum, err)
			}
		}
	}
	if len(rows) > 0 {
		validation := excelize.NewDataValidation(true)
		validation.Sqref = fmt.Sprintf("I2:I%d", len(rows)+1)
		if err := validation.SetRange(0, maxProposedPOUnits, excelize.DataValidationTypeWhole, excelize.DataValidationOperatorBetween); err != nil {
			return fmt.Errorf("failed to configure proposed PO validation: %w", err)
		}
		// The I1 header note explains the input; a selection prompt would obscure
		// adjacent rows while entering proposed units.
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
			{"AND(ISNUMBER(J2),J2<=1)", 0},
			{"AND(ISNUMBER(J2),J2>1,J2<=3)", 2},
			{"AND(ISNUMBER(J2),J2>3)", 4},
		} {
			fill := standardSheetCellFillColor("", mtoColumnIdx, mtoColumnIdx, -1, band.value, 0, band.value)
			style, err := f.NewConditionalStyle(&excelize.Style{Fill: patternFill(fill)})
			if err != nil {
				return fmt.Errorf("failed to create proposed PO MTO color: %w", err)
			}
			rules = append(rules, excelize.ConditionalFormatOptions{Type: "formula", Criteria: band.criteria, Format: &style})
		}
		if err := f.SetConditionalFormat(mtoSheetName, fmt.Sprintf("J2:J%d", len(rows)+1), rules); err != nil {
			return fmt.Errorf("failed to color proposed PO MTO: %w", err)
		}
	}
	// Excelize adds an Excel dropdown to every column header for filtering.
	if err := f.AutoFilter(mtoSheetName, "A1:R1", nil); err != nil {
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

// writeMTOScenarioFormula links rowNum's proposed-PO input to its live Excel MTO
// and stockout date in f. asOf is the report date used for already depleted stock.
// first and last identify this item's reserved helper rows; they are unused when
// row has no usable demand. It returns the first formula-writing error.
func writeMTOScenarioFormula(f *excelize.File, row mtoForecastRow, asOf time.Time, rowNum, first, last int) error {
	quantity := fmt.Sprintf("E%d+I%d", rowNum, rowNum)
	reportDate := fmt.Sprintf("DATE(%d,%d,%d)", asOf.Year(), asOf.Month(), asOf.Day())
	result := fmt.Sprintf("IF(%s<=0,0,\"Insufficient history\")", quantity)
	dateResult := fmt.Sprintf("IF(%s<=0,%s,\"Insufficient history\")", quantity, reportDate)
	if row.hasDemand {
		// ref returns the helper-sheet range for this item's forecast segments.
		ref := func(col string) string {
			return fmt.Sprintf("'%s'!$%s$%d:$%s$%d", mtoScenarioSheetName, col, first, col, last)
		}
		index := fmt.Sprintf("MIN(COUNTIF(%s,\"<\"&(%s))+1,%d)", ref("D"), quantity, last-first+1)
		fraction := fmt.Sprintf("(%s-INDEX(%s,%s))/INDEX(%s,%s)", quantity, ref("C"), index, ref("B"), index)
		// The last cumulative value bounds the simulated horizon.
		result = fmt.Sprintf("IF(%s<=0,0,IF(%s>'%s'!D%d,\">%d\",IFERROR(INDEX(%s,%s)+%s,\">%d\")))",
			quantity, quantity, mtoScenarioSheetName, last, mtoHorizonMonths, ref("E"), index, fraction, mtoHorizonMonths)
		// Helper F holds the segment's actual start date; G holds days in its
		// full calendar month. Do not convert fractional MTO using average days.
		// Round days to the same precision as Go before INT takes the date,
		// avoiding floating-point noise at exact midnight boundaries.
		dateResult = fmt.Sprintf("IF(%s<=0,%s,IF(%s>'%s'!D%d,\"Not within %d months\",IFERROR(INDEX(%s,%s)+INT(ROUND((%s)*INDEX(%s,%s),%d)),\"Not within %d months\")))",
			quantity, reportDate, quantity, mtoScenarioSheetName, last, mtoHorizonMonths, ref("F"), index, fraction, ref("G"), index, mtoStockoutDayDecimals, mtoHorizonMonths)
	}
	valid := fmt.Sprintf("OR(I%d=\"\",IFERROR(AND(ISNUMBER(I%d),I%d>=0,I%d=INT(I%d)),FALSE))", rowNum, rowNum, rowNum, rowNum, rowNum)
	for _, formula := range []struct{ column, result string }{
		{"J", result}, {"K", dateResult},
	} {
		// Excelize stores formulas for Excel recalculation on open; both results
		// share input validation so invalid quantities never appear as dates.
		if err := f.SetCellFormula(mtoSheetName, fmt.Sprintf("%s%d", formula.column, rowNum), fmt.Sprintf("IFERROR(IF(%s,%s,\"Invalid PO units\"),\"Invalid PO units\")", valid, formula.result)); err != nil {
			return fmt.Errorf("failed to set proposed PO formula %s%d: %w", formula.column, rowNum, err)
		}
	}
	return nil
}

// mtoScenarioSegment identifies a calendar month and the fraction of its demand
// included in the forecast horizon, with a start date and full month length for
// daily depletion estimates. All items share these date-only segments.
type mtoScenarioSegment struct {
	start      time.Time
	fraction   float64
	monthDays  int
	monthIndex int
}

// mtoScenarioSegments returns the prorated calendar-month segments from the day
// after asOf through the 24-month horizon, matching the baseline Go simulation.
func mtoScenarioSegments(asOf time.Time) []mtoScenarioSegment {
	start := asOf.AddDate(0, 0, 1)
	end := monthsAfter(start, mtoHorizonMonths)
	segments := make([]mtoScenarioSegment, 0, mtoHorizonMonths+1)
	for monthStart := time.Date(start.Year(), start.Month(), 1, 0, 0, 0, 0, time.UTC); monthStart.Before(end); monthStart = monthStart.AddDate(0, 1, 0) {
		monthEnd := monthStart.AddDate(0, 1, 0)
		segmentStart, segmentEnd := monthStart, monthEnd
		if segmentStart.Before(start) {
			segmentStart = start
		}
		if segmentEnd.After(end) {
			segmentEnd = end
		}
		segments = append(segments, mtoScenarioSegment{
			start:      segmentStart,
			monthDays:  monthEnd.AddDate(0, 0, -1).Day(),
			fraction:   segmentEnd.Sub(segmentStart).Hours() / monthEnd.Sub(monthStart).Hours(),
			monthIndex: int(monthStart.Month()) - 1,
		})
	}
	return segments
}

// writeMTOForecastDataSheet streams rows' monthly forecast segments into f's
// reserved hidden helper tab. rows and asOf must match writeMTOSheet's inputs.
// Call this after all sheet-visibility changes, and save f before reading the
// helper data or calculating dependent formulas. It returns an error for a
// missing report date or a failed worksheet write. Helper columns A–G contain
// SKU, full-month units, cumulative units before/after the segment, elapsed
// months, segment start as an Excel date, and full calendar-month days.
func writeMTOForecastDataSheet(f *excelize.File, rows []mtoForecastRow, asOf time.Time) error {
	if asOf.IsZero() {
		return fmt.Errorf("cannot create MTO forecast data without a sales history run date")
	}
	// Excelize serializes ascending rows without retaining a worksheet object
	// graph, spilling large XML buffers to temporary files as needed.
	stream, err := f.NewStreamWriter(mtoScenarioSheetName)
	if err != nil {
		return fmt.Errorf("failed to stream MTO forecast data: %w", err)
	}
	segments := mtoScenarioSegments(asOf)
	dataRow := mtoScenarioFirstDataRow
	for _, row := range rows {
		if !row.hasDemand {
			continue
		}
		cumulative, elapsed := 0.0, 0.0
		for _, segment := range segments {
			units := row.monthly[segment.monthIndex]
			// Excelize serializes time.Time as a numeric Excel date so formulas
			// can add fractional-month demand directly in calendar days.
			values := [...]interface{}{row.item.SKU, units, cumulative, cumulative + units*segment.fraction, elapsed, segment.start, segment.monthDays}
			if err := stream.SetRow(fmt.Sprintf("A%d", dataRow), values[:]); err != nil {
				return fmt.Errorf("failed to write MTO forecast data row %d: %w", dataRow, err)
			}
			cumulative += units * segment.fraction
			elapsed += segment.fraction
			dataRow++
		}
	}
	// Flush finalizes the worksheet XML and reports temporary-file write errors.
	if err := stream.Flush(); err != nil {
		return fmt.Errorf("failed to flush MTO forecast data: %w", err)
	}
	return nil
}
