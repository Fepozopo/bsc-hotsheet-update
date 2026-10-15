package hotsheet

import (
	"fmt"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	dashboardSheetName     = "Dashboard"
	dashboardDataSheetName = "Dashboard Data"
	dashboardClassList     = "DashboardClasses"
	// Fixed row heights keep the 480px donuts clear of the lower controls,
	// regardless of the spreadsheet application's default font metrics.
	dashboardRowHeight   = 15.0
	dashboardBarChartRow = 38
	// Each class occupies one summary row; rankings store labels then values.
	dashboardShippedLabels = 2
	dashboardShippedValues = dashboardShippedLabels + dashboardTopItems
	dashboardDollarLabels  = dashboardShippedValues + dashboardTopItems
	dashboardDollarValues  = dashboardDollarLabels + dashboardTopItems
	dashboardSalesValues   = dashboardDollarValues + dashboardTopItems
	dashboardMonthlyValues = dashboardSalesValues + 2
)

// writeDashboardSheet fills the first worksheet with independently filtered YTD
// donuts and sales comparisons. hasHistory enables monthly sales and completed-
// month comparisons anchored to runDate. It returns worksheet or chart errors.
func writeDashboardSheet(f *excelize.File, entries []*inventoryEntry, hasHistory bool, runDate time.Time) error {
	data := buildDashboardData(entries, hasHistory, runDate)
	if _, err := f.NewSheet(dashboardDataSheetName); err != nil {
		return err
	}
	for row, summary := range data {
		values := []interface{}{summary.class}
		for _, ranking := range [][dashboardTopItems]dashboardItem{summary.shipped, summary.dollars} {
			for _, item := range ranking {
				values = append(values, item.label)
			}
			for _, item := range ranking {
				values = append(values, item.value)
			}
		}
		values = append(values, summary.sales[0], summary.sales[1])
		for _, year := range summary.monthly {
			for _, amount := range year {
				values = append(values, amount)
			}
		}
		if err := f.SetSheetRow(dashboardDataSheetName, fmt.Sprintf("A%d", row+1), &values); err != nil {
			return err
		}
	}
	// A named range supports cross-sheet validation and class lists longer than
	// Excel's 255-character limit for inline dropdowns.
	if err := f.SetDefinedName(&excelize.DefinedName{Name: dashboardClassList, RefersTo: fmt.Sprintf("'%s'!$A$1:$A$%d", dashboardDataSheetName, len(data))}); err != nil {
		return err
	}
	if err := formatDashboard(f, hasHistory, runDate); err != nil {
		return err
	}

	// Helper tables sit below the summaries, so adding classes cannot overwrite
	// chart formulas. Each table references only its own Dashboard selector.
	firstRow := len(data) + 3
	for _, spec := range []struct {
		selector, anchor, heading, title string
		labels, values, helperCol        int
	}{
		{"B4", "A6", "A5", "Top 15 Items by Quantity Shipped YTD", dashboardShippedLabels, dashboardShippedValues, 1},
		{"L4", "K6", "K5", "Top 15 Items by Dollars Sold YTD", dashboardDollarLabels, dashboardDollarValues, 5},
	} {
		if err := addDashboardSelector(f, spec.selector); err != nil {
			return err
		}
		for rank := range dashboardTopItems {
			if err := writeDashboardLookup(f, firstRow+rank, spec.helperCol, spec.labels+rank, spec.selector, len(data)); err != nil {
				return err
			}
			if err := writeDashboardLookup(f, firstRow+rank, spec.helperCol+1, spec.values+rank, spec.selector, len(data)); err != nil {
				return err
			}
		}
		// Worksheet headings stay top-left and outside the chart's automatic
		// title/legend layout, even when product descriptions wrap over many lines.
		if err := f.SetCellValue(dashboardSheetName, spec.heading, spec.title); err != nil {
			return err
		}
		chart := &excelize.Chart{
			Type: excelize.Doughnut, HoleSize: 60,
			Dimension: excelize.ChartDimension{Width: 700, Height: 480},
			Legend:    excelize.ChartLegend{Position: "r"},
			PlotArea:  excelize.ChartPlotArea{ShowPercent: true},
			Series: []excelize.ChartSeries{{Name: fmt.Sprintf("'%s'!%s", dashboardSheetName, spec.heading),
				Categories: dashboardRange(spec.helperCol, firstRow, dashboardTopItems),
				Values:     dashboardRange(spec.helperCol+1, firstRow, dashboardTopItems)}},
		}
		// Excelize charts include hidden-sheet data (plotVisOnly=false), allowing
		// formulas and summaries to stay out of the user-facing Dashboard.
		if err := f.AddChart(dashboardSheetName, spec.anchor, chart); err != nil {
			return err
		}
	}

	if err := addDashboardSelector(f, "B36"); err != nil {
		return err
	}
	comparisonTitle := "YTD Sales vs Full Previous Year"
	labels := [2]string{"YTD", "Previous Year (Full Year)"}
	if hasHistory {
		comparisonTitle = dashboardCompletedPeriod(runDate) + " Sales vs Previous Year"
		labels = [2]string{fmt.Sprintf("%d YTD (completed months)", runDate.Year()), fmt.Sprintf("%d (same months)", runDate.Year()-1)}
	}
	for year := range 2 {
		if err := f.SetCellValue(dashboardDataSheetName, fmt.Sprintf("M%d", firstRow+year), labels[year]); err != nil {
			return err
		}
		if err := writeDashboardLookup(f, firstRow+year, 14, dashboardSalesValues+year, "B36", len(data)); err != nil {
			return err
		}
	}
	if err := f.SetCellStr(dashboardDataSheetName, fmt.Sprintf("N%d", firstRow-1), "Sales"); err != nil {
		return err
	}
	comparison := dashboardBarChart(comparisonTitle, []excelize.ChartSeries{{Name: fmt.Sprintf("'%s'!$N$%d", dashboardDataSheetName, firstRow-1), Categories: dashboardRange(13, firstRow, 2), Values: dashboardRange(14, firstRow, 2)}})
	comparison.PlotArea.ShowVal = true
	if err := f.AddChart(dashboardSheetName, fmt.Sprintf("A%d", dashboardBarChartRow), comparison); err != nil {
		return err
	}

	if hasHistory {
		if err := addDashboardSelector(f, "L36"); err != nil {
			return err
		}
		for month := range 12 {
			if err := f.SetCellValue(dashboardDataSheetName, fmt.Sprintf("I%d", firstRow+month), time.Month(month + 1).String()[:3]); err != nil {
				return err
			}
			for year := range 2 {
				if err := writeDashboardLookup(f, firstRow+month, 10+year, dashboardMonthlyValues+year*12+month, "L36", len(data)); err != nil {
					return err
				}
			}
		}
		series := make([]excelize.ChartSeries, 2)
		for year := range series {
			cell, _ := excelize.CoordinatesToCellName(10+year, firstRow-1)
			// Excelize serializes Name as a reference, not a literal label.
			// Store the year as text to supply a valid Excel legend reference.
			if err := f.SetCellStr(dashboardDataSheetName, cell, strconv.Itoa(runDate.Year()-year)); err != nil {
				return err
			}
			series[year] = excelize.ChartSeries{Name: fmt.Sprintf("'%s'!%s", dashboardDataSheetName, cell), Categories: dashboardRange(9, firstRow, 12), Values: dashboardRange(10+year, firstRow, 12)}
		}
		if err := f.AddChart(dashboardSheetName, fmt.Sprintf("K%d", dashboardBarChartRow), dashboardBarChart("Monthly Sales: This Year vs Last Year", series)); err != nil {
			return err
		}
	}
	// Recalculate class-lookup formulas whenever a selector changes, including
	// after opening the generated file in Excel; no macros or refresh are needed.
	mode, recalc := "auto", true
	if err := f.SetCalcProps(&excelize.CalcPropsOptions{CalcMode: &mode, FullCalcOnLoad: &recalc, ForceFullCalc: &recalc}); err != nil {
		return err
	}
	return f.SetSheetVisible(dashboardDataSheetName, false)
}

// writeDashboardLookup writes a class-sensitive INDEX/MATCH formula into a helper
// cell. Source columns identify summary metrics; worksheet errors are returned.
// Wildcard escaping makes class names containing *, ? or ~ literal. Excel computes
// the initial values on open using the workbook's automatic calculation settings.
func writeDashboardLookup(f *excelize.File, row, col, sourceCol int, selector string, classCount int) error {
	cell, _ := excelize.CoordinatesToCellName(col, row)
	column, _ := excelize.ColumnNumberToName(sourceCol)
	selected := fmt.Sprintf("'%s'!%s", dashboardSheetName, selector)
	literal := fmt.Sprintf(`SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(%s,"~","~~"),"*","~*"),"?","~?")`, selected)
	formula := fmt.Sprintf("INDEX($%s$1:$%s$%d,MATCH(%s,$A$1:$A$%d,0),1)", column, column, classCount, literal, classCount)
	return f.SetCellFormula(dashboardDataSheetName, cell, formula)
}

// dashboardRange returns an absolute hidden-sheet column range for count rows.
func dashboardRange(col, firstRow, count int) string {
	column, _ := excelize.ColumnNumberToName(col)
	return fmt.Sprintf("'%s'!$%s$%d:$%s$%d", dashboardDataSheetName, column, firstRow, column, firstRow+count-1)
}

// addDashboardSelector writes All and a strict class-list validation at cell;
// Excel's validation supplies the dropdown and rejects unrecognized class names.
func addDashboardSelector(f *excelize.File, cell string) error {
	if err := f.SetCellValue(dashboardSheetName, cell, "All"); err != nil {
		return err
	}
	validation := excelize.NewDataValidation(false)
	validation.SetSqref(cell)
	validation.SetSqrefDropList(dashboardClassList)
	validation.SetError(excelize.DataValidationErrorStyleStop, "Choose a class", "Select All or a class from the dropdown.")
	return f.AddDataValidation(dashboardSheetName, validation)
}

// dashboardBarChart returns a clustered column chart with dollar-formatted sales
// and the supplied series. Two monthly series produce adjacent bars per month.
func dashboardBarChart(title string, series []excelize.ChartSeries) *excelize.Chart {
	return &excelize.Chart{
		Type: excelize.Col, Series: series,
		Title:     []excelize.RichTextRun{{Text: title}},
		Dimension: excelize.ChartDimension{Width: 700, Height: 420},
		Legend:    excelize.ChartLegend{Position: "b"},
		YAxis:     excelize.ChartAxis{NumFmt: excelize.ChartNumFmt{CustomNumFmt: "$#,##0"}, MajorGridLines: true},
	}
}

// dashboardCompletedPeriod describes the current year's completed-month window;
// January explicitly has no completed months rather than using last year's YTD.
func dashboardCompletedPeriod(runDate time.Time) string {
	if runDate.Month() == time.January {
		return fmt.Sprintf("%d: No Completed Months", runDate.Year())
	}
	return fmt.Sprintf("Jan–%s %d", runDate.Month()-1, runDate.Year())
}

// formatDashboard writes explanatory period labels and lays out chart controls.
// hasHistory selects the reporting notes. Donut headings have their own left-
// aligned row, and fixed row heights reserve space above the lower controls.
// Style and worksheet errors are returned.
func formatDashboard(f *excelize.File, hasHistory bool, runDate time.Time) error {
	notes := "Donuts: inventory YTD; positive values only. Percentages are shares of the top 15, not all sales. Each class filter is independent."
	period := "Sales comparison: inventory YTD vs full previous-year dollars sold. Monthly history was not supplied."
	if hasHistory {
		period = fmt.Sprintf("Sales comparison: %s vs the same months in %d. Monthly chart includes Jan–Dec; current month may be partial. Report: %s.", dashboardCompletedPeriod(runDate), runDate.Year()-1, runDate.Format("Jan 2, 2006"))
	}
	cells := map[string]string{"A1": "Dashboard", "A2": notes, "A3": period, "A4": "Class:", "K4": "Class:", "A36": "Class:"}
	if hasHistory {
		cells["K36"] = "Class:"
	}
	for cell, text := range cells {
		if err := f.SetCellValue(dashboardSheetName, cell, text); err != nil {
			return err
		}
	}
	for _, row := range []int{2, 3} {
		if err := f.MergeCell(dashboardSheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("T%d", row)); err != nil {
			return err
		}
	}
	for row := 1; row <= dashboardBarChartRow; row++ {
		if err := f.SetRowHeight(dashboardSheetName, row, dashboardRowHeight); err != nil {
			return err
		}
	}
	if err := f.SetColWidth(dashboardSheetName, "A", "T", 10); err != nil {
		return err
	}
	for _, span := range [][2]string{{"B4", "E4"}, {"L4", "O4"}, {"B36", "E36"}, {"A5", "J5"}, {"K5", "T5"}} {
		if err := f.MergeCell(dashboardSheetName, span[0], span[1]); err != nil {
			return err
		}
	}
	if hasHistory {
		if err := f.MergeCell(dashboardSheetName, "L36", "O36"); err != nil {
			return err
		}
	}
	heading, err := f.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true, Size: 20}})
	if err != nil {
		return err
	}
	if err := f.SetCellStyle(dashboardSheetName, "A1", "A1", heading); err != nil {
		return err
	}
	donutHeading, err := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 14},
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
	})
	if err != nil {
		return err
	}
	for _, cell := range []string{"A5", "K5"} {
		if err := f.SetCellStyle(dashboardSheetName, cell, cell, donutHeading); err != nil {
			return err
		}
	}
	if err := f.SetRowHeight(dashboardSheetName, 5, 24); err != nil {
		return err
	}
	control, err := f.NewStyle(&excelize.Style{Fill: patternFill(standardHeaderFill), Font: boldFont()})
	if err != nil {
		return err
	}
	for _, cell := range []string{"B4", "L4", "B36"} {
		if err := f.SetCellStyle(dashboardSheetName, cell, cell, control); err != nil {
			return err
		}
	}
	if hasHistory {
		if err := f.SetCellStyle(dashboardSheetName, "L36", "L36", control); err != nil {
			return err
		}
	}
	grid := false
	return f.SetSheetView(dashboardSheetName, 0, &excelize.ViewOptions{ShowGridLines: &grid})
}
