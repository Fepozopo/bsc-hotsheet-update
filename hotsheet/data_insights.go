package hotsheet

import (
	"fmt"
	"math"
	"sort"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	dataInsightsSheetName             = "Data Insights"
	dataInsightsTitleText             = "Data Insights"
	dataInsightsLeftTableStartCol     = "A"
	dataInsightsRightTableStartCol    = "G"
	dataInsightsTitleRow              = 1
	dataInsightsSectionRow            = 3
	dataInsightsTableStartRow         = 5
	dataInsightsTableColumnCount      = 5
	dataInsightsOtherTableColumnCount = 6
)

const (
	dataInsightsColumnOccasion = iota
	dataInsightsColumnDate
	dataInsightsColumnYTD
	dataInsightsColumnPY
	dataInsightsColumnFinal
)

var dataInsightsTableColumnWidths = []float64{26, 31, 14, 14, 34}
var dataInsightsOtherTableColumnWidths = []float64{26, 24, 31, 14, 14, 34}

// dataInsightsCell constructs an Excel cell name (e.g. "A5") from a column letter and row number.
func dataInsightsCell(col string, row int) string {
	return fmt.Sprintf("%s%d", col, row)
}

// dataInsightsTableEndCol calculates the last column of a Data Insights table based on the shared column count.
func dataInsightsTableEndCol(startCol string, columnCount int) (string, error) {
	cols, err := dataInsightsTableColumns(startCol, columnCount)
	if err != nil {
		return "", err
	}

	return cols[columnCount-1], nil
}

// dataInsightsRow represents one output row in the Data Insights sheet.
type dataInsightsRow struct {
	Class           string
	Occasion        string
	Date            string
	DollarSoldYTD   float64
	DollarSoldPY    float64
	ProjectedDollar float64
	Final           string
	sortKey         int
	sortOccasion    string
}

// dataInsightsGroup aggregates sales for a single normalized occasion within a section.
type dataInsightsGroup struct {
	Section             string
	Class               string
	Occasion            string
	Date                string
	sortKey             int
	complete            bool
	TargetMonthsThrough float64
	DollarSoldYTD       float64
	DollarSoldPY        float64
}

// occasionDateInfo captures how an occasion should be displayed, ordered, and projected.
type occasionDateInfo struct {
	Display             string
	Month               time.Month
	Day                 int
	SortKey             int
	Complete            bool
	TargetMonthsThrough float64
}

// dataInsightDateMap maps normalized occasion names to their display text and calendar metadata.
var dataInsightDateMap = map[string]occasionDateInfo{
	// Valentine's Day cards sell in two separate windows, but we still anchor the occasion to Feb 14
	// so the event keeps its normal calendar identity while projection logic handles the split season.
	"VALENTINE'S DAY":   {Display: "February 14", Month: time.February, Day: 14, SortKey: 214},
	"VALENTINES DAY":    {Display: "February 14", Month: time.February, Day: 14, SortKey: 214},
	"ST PATRICKS DAY":   {Display: "March 17", Month: time.March, Day: 17, SortKey: 317},
	"ST. PATRICK'S DAY": {Display: "March 17", Month: time.March, Day: 17, SortKey: 317},
	"EASTER":            {Display: "April 20", Month: time.April, Day: 20, SortKey: 420},
	"MOTHER'S DAY":      {Display: "May 11", Month: time.May, Day: 11, SortKey: 511},
	"MOTHERS DAY":       {Display: "May 11", Month: time.May, Day: 11, SortKey: 511},
	"GRADUATION":        {Display: "mid-June", Month: time.June, Day: 15, SortKey: 615},
	"FATHER'S DAY":      {Display: "June 15", Month: time.June, Day: 15, SortKey: 615},
	"FATHERS DAY":       {Display: "June 15", Month: time.June, Day: 15, SortKey: 615},
	"INDEPENDENCE DAY":  {Display: "July 4", Month: time.July, Day: 4, SortKey: 704},
	"HALLOWEEN":         {Display: "October 31", Month: time.October, Day: 31, SortKey: 1031},
	"VETERAN'S DAY":     {Display: "November 11", Month: time.November, Day: 11, SortKey: 1111},
	"VETERANS DAY":      {Display: "November 11", Month: time.November, Day: 11, SortKey: 1111},
	"THANKSGIVING":      {Display: "November 28", Month: time.November, Day: 28, SortKey: 1128},
	"HANUKKAH":          {Display: "December 8", Month: time.December, Day: 8, SortKey: 1208},
	"HOLIDAY":           {Display: "December 25", Month: time.December, Day: 25, SortKey: 1225},
	"CHRISTMAS":         {Display: "December 25", Month: time.December, Day: 25, SortKey: 1225},
}

// writeDataInsightsSheet creates the "Data Insights" worksheet and populates it with grouped sales data.
func writeDataInsightsSheet(f *excelize.File, entries []*entry) error {
	now := time.Now()
	currentMonthsThrough := currentMonthsThrough(now)
	// Use the current month progress to annualize in-progress rows.
	rowsBySection := buildDataInsightsRows(entries, currentMonthsThrough, now)
	otherRowsBySection := buildOtherProductDataInsightsRows(entries, currentMonthsThrough, now)

	sheetName := dataInsightsSheetName
	if _, err := f.NewSheet(sheetName); err != nil {
		return fmt.Errorf("failed to create %s sheet: %w", sheetName, err)
	}

	titleStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(),
		Font:      boldFont(14),
	})
	if err != nil {
		return fmt.Errorf("failed to create title style: %w", err)
	}

	sectionStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(),
		Border:    thinBlackBorder(),
		Fill:      patternFill(dataInsightsSectionFill),
		Font:      boldFont(),
	})
	if err != nil {
		return fmt.Errorf("failed to create section style: %w", err)
	}

	headerStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(),
		Border:    thinBlackBorder(),
		Fill:      patternFill(standardHeaderFill),
		Font:      boldFont(),
	})
	if err != nil {
		return fmt.Errorf("failed to create header style: %w", err)
	}

	dataStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(),
		Border:    thinBlackBorder(),
	})
	if err != nil {
		return fmt.Errorf("failed to create data style: %w", err)
	}

	currencyDataStyle, err := f.NewStyle(&excelize.Style{
		Alignment:    centeredAlignment(),
		Border:       thinBlackBorder(),
		CustomNumFmt: currencyNumFmt(),
	})
	if err != nil {
		return fmt.Errorf("failed to create currency data style: %w", err)
	}

	totalStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(),
		Border:    thinBlackBorder(),
		Fill:      patternFill(standardTotalFill),
		Font:      boldFont(),
	})
	if err != nil {
		return fmt.Errorf("failed to create total style: %w", err)
	}

	currencyTotalStyle, err := f.NewStyle(&excelize.Style{
		Alignment:    centeredAlignment(),
		Border:       thinBlackBorder(),
		Fill:         patternFill(standardTotalFill),
		Font:         boldFont(),
		CustomNumFmt: currencyNumFmt(),
	})
	if err != nil {
		return fmt.Errorf("failed to create currency total style: %w", err)
	}

	if err := setDataInsightsTableWidths(f, sheetName, dataInsightsLeftTableStartCol, dataInsightsTableColumnWidths); err != nil {
		return err
	}
	if err := setDataInsightsTableWidths(f, sheetName, dataInsightsRightTableStartCol, dataInsightsOtherTableColumnWidths); err != nil {
		return err
	}

	leftEndCol, err := dataInsightsTableEndCol(dataInsightsLeftTableStartCol, dataInsightsTableColumnCount)
	if err != nil {
		return err
	}
	titleEndCol, err := dataInsightsTableEndCol(dataInsightsRightTableStartCol, dataInsightsOtherTableColumnCount)
	if err != nil {
		return err
	}

	// The title spans the full width of both tables to stay visually centered on the sheet.
	titleCell := dataInsightsCell(dataInsightsLeftTableStartCol, dataInsightsTitleRow)
	if err := f.SetCellValue(sheetName, titleCell, dataInsightsTitleText); err != nil {
		return fmt.Errorf("failed to set sheet title: %w", err)
	}
	if err := f.MergeCell(sheetName, titleCell, dataInsightsCell(titleEndCol, dataInsightsTitleRow)); err != nil {
		return fmt.Errorf("failed to merge title cells: %w", err)
	}
	if err := f.SetCellStyle(sheetName, titleCell, dataInsightsCell(titleEndCol, dataInsightsTitleRow), titleStyle); err != nil {
		return fmt.Errorf("failed to style title row: %w", err)
	}

	// The "Counter Cards" subtitle for the left-side table.
	counterCardsCell := dataInsightsCell(dataInsightsLeftTableStartCol, dataInsightsSectionRow)
	if err := f.SetCellValue(sheetName, counterCardsCell, "Counter Cards"); err != nil {
		return fmt.Errorf("failed to set counter cards subtitle: %w", err)
	}
	if err := f.MergeCell(sheetName, counterCardsCell, dataInsightsCell(leftEndCol, dataInsightsSectionRow)); err != nil {
		return fmt.Errorf("failed to merge counter cards subtitle: %w", err)
	}
	if err := f.SetCellStyle(sheetName, counterCardsCell, dataInsightsCell(leftEndCol, dataInsightsSectionRow), sectionStyle); err != nil {
		return fmt.Errorf("failed to style counter cards subtitle: %w", err)
	}

	// The "Other Products" subtitle for the right side table.
	otherProductsCell := dataInsightsCell(dataInsightsRightTableStartCol, dataInsightsSectionRow)
	if err := f.SetCellValue(sheetName, otherProductsCell, "Other Products"); err != nil {
		return fmt.Errorf("failed to set other products subtitle: %w", err)
	}
	if err := f.MergeCell(sheetName, otherProductsCell, dataInsightsCell(titleEndCol, dataInsightsSectionRow)); err != nil {
		return fmt.Errorf("failed to merge other products subtitle: %w", err)
	}
	if err := f.SetCellStyle(sheetName, otherProductsCell, dataInsightsCell(titleEndCol, dataInsightsSectionRow), sectionStyle); err != nil {
		return fmt.Errorf("failed to style other products subtitle: %w", err)
	}

	leftSections := []dataInsightsSection{
		{
			Name:               "Spring",
			Headers:            []string{"Occasion", "Date", "YTD Sales", "PY Sales", "Status / Projected YoY"},
			Rows:               rowsBySection["Spring"],
			CurrencyStartIndex: dataInsightsColumnYTD,
			CurrencyEndIndex:   dataInsightsColumnPY,
			RenderRow: func(row dataInsightsRow) []interface{} {
				return []interface{}{row.Occasion, row.Date, row.DollarSoldYTD, row.DollarSoldPY, row.Final}
			},
			RenderTotal: func(totalYTD, totalPY, totalProjected float64, rows []dataInsightsRow) []interface{} {
				return []interface{}{"Total", "", totalYTD, totalPY, dataInsightsSeasonTotalFinal(totalProjected, totalPY, rows)}
			},
		},
		{
			Name:               "Winter",
			Headers:            []string{"Occasion", "Date", "YTD Sales", "PY Sales", "Status / Projected YoY"},
			Rows:               rowsBySection["Winter"],
			CurrencyStartIndex: dataInsightsColumnYTD,
			CurrencyEndIndex:   dataInsightsColumnPY,
			RenderRow: func(row dataInsightsRow) []interface{} {
				return []interface{}{row.Occasion, row.Date, row.DollarSoldYTD, row.DollarSoldPY, row.Final}
			},
			RenderTotal: func(totalYTD, totalPY, totalProjected float64, rows []dataInsightsRow) []interface{} {
				return []interface{}{"Total", "", totalYTD, totalPY, dataInsightsSeasonTotalFinal(totalProjected, totalPY, rows)}
			},
		},
		{
			Name:               "Everyday",
			Headers:            []string{"Occasion", "Date", "YTD Sales", "PY Sales", "Projected YoY"},
			Rows:               rowsBySection["Everyday"],
			CurrencyStartIndex: dataInsightsColumnYTD,
			CurrencyEndIndex:   dataInsightsColumnPY,
			RenderRow: func(row dataInsightsRow) []interface{} {
				return []interface{}{row.Occasion, row.Date, row.DollarSoldYTD, row.DollarSoldPY, row.Final}
			},
			RenderTotal: func(totalYTD, totalPY, totalProjected float64, rows []dataInsightsRow) []interface{} {
				return []interface{}{"Total", "", totalYTD, totalPY, formatYoYFromProjectedSales(totalProjected, totalPY)}
			},
		},
	}

	rowNum := dataInsightsTableStartRow
	for idx, section := range leftSections {
		if idx > 0 {
			rowNum++
		}
		nextRow, err := writeDataInsightsSectionTable(f, sheetName, dataInsightsLeftTableStartCol, rowNum, section, sectionStyle, headerStyle, dataStyle, currencyDataStyle, totalStyle, currencyTotalStyle)
		if err != nil {
			return err
		}
		rowNum = nextRow
	}

	rightSections := []dataInsightsSection{
		{
			Name:               "Spring",
			Headers:            []string{"Class", "Occasion", "Date", "YTD Sales", "PY Sales", "Status / Projected YoY"},
			Rows:               otherRowsBySection["Spring"],
			CurrencyStartIndex: 3,
			CurrencyEndIndex:   4,
			RenderRow: func(row dataInsightsRow) []interface{} {
				return []interface{}{row.Class, row.Occasion, row.Date, row.DollarSoldYTD, row.DollarSoldPY, row.Final}
			},
			RenderTotal: func(totalYTD, totalPY, totalProjected float64, rows []dataInsightsRow) []interface{} {
				return []interface{}{"Total", "", "", totalYTD, totalPY, dataInsightsSeasonTotalFinal(totalProjected, totalPY, rows)}
			},
		},
		{
			Name:               "Winter",
			Headers:            []string{"Class", "Occasion", "Date", "YTD Sales", "PY Sales", "Status / Projected YoY"},
			Rows:               otherRowsBySection["Winter"],
			CurrencyStartIndex: 3,
			CurrencyEndIndex:   4,
			RenderRow: func(row dataInsightsRow) []interface{} {
				return []interface{}{row.Class, row.Occasion, row.Date, row.DollarSoldYTD, row.DollarSoldPY, row.Final}
			},
			RenderTotal: func(totalYTD, totalPY, totalProjected float64, rows []dataInsightsRow) []interface{} {
				return []interface{}{"Total", "", "", totalYTD, totalPY, dataInsightsSeasonTotalFinal(totalProjected, totalPY, rows)}
			},
		},
		{
			Name:               "Everyday",
			Headers:            []string{"Class", "Occasion", "Date", "YTD Sales", "PY Sales", "Projected YoY"},
			Rows:               otherRowsBySection["Everyday"],
			CurrencyStartIndex: 3,
			CurrencyEndIndex:   4,
			RenderRow: func(row dataInsightsRow) []interface{} {
				return []interface{}{row.Class, row.Occasion, row.Date, row.DollarSoldYTD, row.DollarSoldPY, row.Final}
			},
			RenderTotal: func(totalYTD, totalPY, totalProjected float64, rows []dataInsightsRow) []interface{} {
				return []interface{}{"Total", "", "", totalYTD, totalPY, formatYoYFromProjectedSales(totalProjected, totalPY)}
			},
		},
	}

	rowNum = dataInsightsTableStartRow
	for idx, section := range rightSections {
		if idx > 0 {
			rowNum++
		}
		nextRow, err := writeDataInsightsSectionTable(f, sheetName, dataInsightsRightTableStartCol, rowNum, section, sectionStyle, headerStyle, dataStyle, currencyDataStyle, totalStyle, currencyTotalStyle)
		if err != nil {
			return err
		}
		rowNum = nextRow
	}

	return nil
}

// dataInsightsSection captures all the information needed to render one section of the Data Insights sheet, including
// the section name, headers, rows, which columns are currency values, and how to render the row and total values.
type dataInsightsSection struct {
	Name               string
	Headers            []string
	Rows               []dataInsightsRow
	CurrencyStartIndex int
	CurrencyEndIndex   int
	RenderRow          func(dataInsightsRow) []interface{}
	RenderTotal        func(totalYTD, totalPY, totalProjected float64, rows []dataInsightsRow) []interface{}
}

// writeDataInsightsSectionTable renders one seasonal section, including the section title,
// headers, detail rows, and total row. The caller supplies the row formatter so the same
// layout code can be reused for both Counter Cards and Other Products.
func writeDataInsightsSectionTable(f *excelize.File, sheetName, startCol string, startRow int, section dataInsightsSection, sectionStyle, headerStyle, dataStyle, currencyDataStyle, totalStyle, currencyTotalStyle int) (int, error) {
	cols, err := dataInsightsTableColumns(startCol, len(section.Headers))
	if err != nil {
		return 0, err
	}
	endCol := cols[len(cols)-1]

	sectionTitleCell := dataInsightsCell(startCol, startRow)
	sectionTitleEnd := dataInsightsCell(endCol, startRow)
	if err := f.SetCellValue(sheetName, sectionTitleCell, section.Name); err != nil {
		return 0, fmt.Errorf("failed to set section title %s: %w", section.Name, err)
	}
	if err := f.MergeCell(sheetName, sectionTitleCell, sectionTitleEnd); err != nil {
		return 0, fmt.Errorf("failed to merge section title %s: %w", section.Name, err)
	}
	if err := f.SetCellStyle(sheetName, sectionTitleCell, sectionTitleEnd, sectionStyle); err != nil {
		return 0, fmt.Errorf("failed to style section title %s: %w", section.Name, err)
	}

	rowNum := startRow + 1
	for colIdx, header := range section.Headers {
		cell := dataInsightsCell(cols[colIdx], rowNum)
		if err := f.SetCellValue(sheetName, cell, header); err != nil {
			return 0, fmt.Errorf("failed to set header %s: %w", header, err)
		}
	}
	if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[0], rowNum), dataInsightsCell(endCol, rowNum), headerStyle); err != nil {
		return 0, fmt.Errorf("failed to style header row for %s: %w", section.Name, err)
	}

	rowNum++
	totalYTD := 0.0
	totalPY := 0.0
	totalProjectedSales := 0.0
	for _, row := range section.Rows {
		values := section.RenderRow(row)
		if len(values) != len(section.Headers) {
			return 0, fmt.Errorf("failed to render %s row: got %d values, want %d", section.Name, len(values), len(section.Headers))
		}
		for colIdx, value := range values {
			cell := dataInsightsCell(cols[colIdx], rowNum)
			if err := f.SetCellValue(sheetName, cell, value); err != nil {
				return 0, fmt.Errorf("failed to write %s row cell %s: %w", section.Name, cell, err)
			}
		}
		if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[0], rowNum), dataInsightsCell(endCol, rowNum), dataStyle); err != nil {
			return 0, fmt.Errorf("failed to style %s data row: %w", section.Name, err)
		}
		if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[section.CurrencyStartIndex], rowNum), dataInsightsCell(cols[section.CurrencyEndIndex], rowNum), currencyDataStyle); err != nil {
			return 0, fmt.Errorf("failed to style %s currency cells: %w", section.Name, err)
		}

		totalYTD += row.DollarSoldYTD
		totalPY += row.DollarSoldPY
		totalProjectedSales += row.ProjectedDollar
		rowNum++
	}

	// The total row keeps the actual YTD and PY sums in their respective columns; only the
	// rightmost column uses projected sales to derive the YoY percentage shown in the sheet.
	totalRowValues := section.RenderTotal(totalYTD, totalPY, totalProjectedSales, section.Rows)
	if len(totalRowValues) != len(section.Headers) {
		return 0, fmt.Errorf("failed to render %s total row: got %d values, want %d", section.Name, len(totalRowValues), len(section.Headers))
	}
	for colIdx, value := range totalRowValues {
		cell := dataInsightsCell(cols[colIdx], rowNum)
		if err := f.SetCellValue(sheetName, cell, value); err != nil {
			return 0, fmt.Errorf("failed to write %s total row cell %s: %w", section.Name, cell, err)
		}
	}
	if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[0], rowNum), dataInsightsCell(endCol, rowNum), totalStyle); err != nil {
		return 0, fmt.Errorf("failed to style total row for %s: %w", section.Name, err)
	}
	if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[section.CurrencyStartIndex], rowNum), dataInsightsCell(cols[section.CurrencyEndIndex], rowNum), currencyTotalStyle); err != nil {
		return 0, fmt.Errorf("failed to style %s total currency cells: %w", section.Name, err)
	}

	return rowNum + 1, nil
}

// dataInsightsSeasonTotalFinal derives the section total status text from the rows that
// were actually written so the total matches the same in-progress/complete wording used by
// the detail rows above it.
func dataInsightsSeasonTotalFinal(totalProjected, totalPY float64, rows []dataInsightsRow) string {
	sectionStarted := true
	sectionComplete := true
	for _, row := range rows {
		if strings.HasPrefix(row.Final, "NOT STARTED") {
			sectionStarted = false
			sectionComplete = false
		}
		if strings.HasPrefix(row.Final, "IN PROGRESS:") {
			sectionComplete = false
		}
	}
	if !sectionStarted {
		return formatSeasonStatusYoY(totalProjected, totalPY, false, false)
	}
	return formatSeasonStatusYoY(totalProjected, totalPY, true, sectionComplete)
}

// buildDataInsightsRows groups Counter Cards into the seasonal sections used by the
// Data Insights sheet, preserving the existing holiday/date/projection rules.
func buildDataInsightsRows(entries []*entry, currentMonthsThrough float64, now time.Time) map[string][]dataInsightsRow {
	groups := make(map[string]*dataInsightsGroup)

	for _, e := range entries {
		// This sheet only tracks exact Counter Cards entries.
		if !isExactCounterCards(e) {
			continue
		}

		// Normalize occasion names so common variants collapse into the same grouped row.
		occasion := normalizeDataInsightsOccasion(e.Occasion)
		section := mapOccasion(occasion)
		dateInfo := dataInsightDateInfo(section, occasion, now)
		// Use the normalized occasion plus section so variants collapse into one rollup bucket.
		groupKey := section + "|" + strings.ToUpper(occasion)

		group, ok := groups[groupKey]
		if !ok {
			group = &dataInsightsGroup{
				Section:             section,
				Occasion:            occasion,
				Date:                dateInfo.Display,
				sortKey:             dateInfo.SortKey,
				complete:            dateInfo.Complete,
				TargetMonthsThrough: dateInfo.TargetMonthsThrough,
			}
			groups[groupKey] = group
		}

		group.DollarSoldYTD += e.DollarSoldYTD
		group.DollarSoldPY += e.DollarSoldPY
	}

	rowsBySection := newDataInsightsRowsBySection()
	for _, group := range groups {
		dateInfo, projected, final := projectDataInsightsRow(group.Section, group.Occasion, group.DollarSoldYTD, group.DollarSoldPY, currentMonthsThrough, now)
		row := dataInsightsRow{
			Occasion:        group.Occasion,
			Date:            dateInfo.Display,
			DollarSoldYTD:   group.DollarSoldYTD,
			DollarSoldPY:    group.DollarSoldPY,
			ProjectedDollar: projected,
			Final:           final,
			sortKey:         group.sortKey,
			sortOccasion:    strings.ToUpper(group.Occasion),
		}
		if group.Section == "Spring" {
			rowsBySection["Spring"] = append(rowsBySection["Spring"], row)
			continue
		}
		if group.Section == "Winter" {
			rowsBySection["Winter"] = append(rowsBySection["Winter"], row)
			continue
		}
		rowsBySection["Everyday"] = append(rowsBySection["Everyday"], row)
	}

	// Sort seasonal rows by calendar order first, then alphabetically for stable ties.
	sortDataInsightsRows(rowsBySection["Spring"], true, false)
	sortDataInsightsRows(rowsBySection["Winter"], true, false)
	sortDataInsightsRows(rowsBySection["Everyday"], false, false)

	return rowsBySection
}

// buildOtherProductDataInsightsRows groups non-card products by class and occasion, then
// applies the same seasonal bucketing and projection rules used by the card section.
func buildOtherProductDataInsightsRows(entries []*entry, currentMonthsThrough float64, now time.Time) map[string][]dataInsightsRow {
	groups := make(map[string]*dataInsightsGroup)

	for _, e := range entries {
		if isExactCounterCards(e) {
			continue
		}

		classDesc := normalizeDataInsightsClassDescription(e)
		occasion := normalizeDataInsightsOccasion(e.Occasion)
		section := mapOccasion(occasion)
		dateInfo := dataInsightDateInfo(section, occasion, now)
		// Keep class and occasion in the grouping key so the same class can appear once per
		// holiday date instead of collapsing all winter merchandise into a single row.
		groupKey := section + "|" + strings.ToUpper(classDesc) + "|" + strings.ToUpper(occasion)

		group, ok := groups[groupKey]
		if !ok {
			group = &dataInsightsGroup{
				Section:             section,
				Class:               classDesc,
				Occasion:            occasion,
				Date:                dateInfo.Display,
				sortKey:             dateInfo.SortKey,
				complete:            dateInfo.Complete,
				TargetMonthsThrough: dateInfo.TargetMonthsThrough,
			}
			groups[groupKey] = group
		}

		group.DollarSoldYTD += e.DollarSoldYTD
		group.DollarSoldPY += e.DollarSoldPY
	}

	rowsBySection := newDataInsightsRowsBySection()
	for _, group := range groups {
		dateInfo, projected, final := projectDataInsightsRow(group.Section, group.Occasion, group.DollarSoldYTD, group.DollarSoldPY, currentMonthsThrough, now)
		row := dataInsightsRow{
			Class:           group.Class,
			Occasion:        group.Occasion,
			Date:            dateInfo.Display,
			DollarSoldYTD:   group.DollarSoldYTD,
			DollarSoldPY:    group.DollarSoldPY,
			ProjectedDollar: projected,
			Final:           final,
			sortKey:         group.sortKey,
			sortOccasion:    strings.ToUpper(group.Occasion),
		}
		rowsBySection[group.Section] = append(rowsBySection[group.Section], row)
	}

	sortDataInsightsRows(rowsBySection["Spring"], true, true)
	sortDataInsightsRows(rowsBySection["Winter"], true, true)
	sortDataInsightsRows(rowsBySection["Everyday"], false, true)

	return rowsBySection
}

// projectDataInsightsRow centralizes the seasonal projection rules so both card and
// non-card Data Insights rows use the same date metadata and year-over-year logic.
func projectDataInsightsRow(section, occasion string, dollarSoldYTD, dollarSoldPY, currentMonthsThrough float64, now time.Time) (occasionDateInfo, float64, string) {
	dateInfo := dataInsightDateInfo(section, occasion, now)

	if section == "Spring" || section == "Winter" {
		// Seasonal items are projected only up to their typical selling window.
		if isValentinesOccasion(occasion) {
			// Valentine's Day is the one split-season exception: it sells in an early-year
			// window and again near the end of the year, so we project against both windows.
			currentSellingDays, totalSellingDays, complete := valentinesProjectionWindow(now)
			projected := dollarSoldYTD * (totalSellingDays / currentSellingDays)
			return dateInfo, projected, formatSeasonStatusYoY(projected, dollarSoldPY, true, complete)
		}
		if section == "Winter" {
			// Winter items are treated as selling from July 1 forward, which matches the
			// existing card logic and avoids projecting winter holidays from the off-season.
			seasonStart := time.Date(now.Year(), time.July, 1, 0, 0, 0, 0, now.Location())
			if now.Before(seasonStart) {
				projected := dollarSoldYTD
				return dateInfo, projected, formatSeasonStatusYoY(projected, dollarSoldPY, false, dateInfo.Complete)
			}
			if dateInfo.Complete {
				projected := dollarSoldYTD
				return dateInfo, projected, formatSeasonStatusYoY(projected, dollarSoldPY, true, dateInfo.Complete)
			}
			currentSellingMonths := monthsThroughSinceDate(now.Year(), time.July, 1, now.Month(), now.Day(), now.Location())
			projected := dollarSoldYTD * (dateInfo.TargetMonthsThrough / currentSellingMonths)
			return dateInfo, projected, formatSeasonStatusYoY(projected, dollarSoldPY, true, dateInfo.Complete)
		}
		if dateInfo.Complete {
			// Once the holiday has passed, stop extrapolating and keep the row at actual YTD.
			projected := dollarSoldYTD
			return dateInfo, projected, formatSeasonStatusYoY(projected, dollarSoldPY, true, dateInfo.Complete)
		}
		projected := dollarSoldYTD * (dateInfo.TargetMonthsThrough / currentMonthsThrough)
		return dateInfo, projected, formatSeasonStatusYoY(projected, dollarSoldPY, true, dateInfo.Complete)
	}

	// Everyday items are projected across the full year because they do not have a holiday
	// cutoff date to anchor them to.
	projected := dollarSoldYTD * (12.0 / currentMonthsThrough)
	return dateInfo, projected, formatYoYFromProjectedSales(projected, dollarSoldPY)
}

// newDataInsightsRowsBySection pre-creates each section so the worksheet keeps a stable
// Spring / Winter / Everyday layout even when one of the buckets is empty.
func newDataInsightsRowsBySection() map[string][]dataInsightsRow {
	return map[string][]dataInsightsRow{
		"Spring":   {},
		"Winter":   {},
		"Everyday": {},
	}
}

// sortDataInsightsRows keeps section ordering stable. Seasonal rows are sorted by holiday
// date first, while Other Products also groups by class so identical classes stay together.
// Within a class, rows then sort by holiday date and occasion label.
func sortDataInsightsRows(rows []dataInsightsRow, useSortKey bool, useClass bool) {
	sort.Slice(rows, func(i, j int) bool {
		if useClass && rows[i].Class != rows[j].Class {
			return strings.ToUpper(rows[i].Class) < strings.ToUpper(rows[j].Class)
		}
		if useSortKey && rows[i].sortKey != rows[j].sortKey {
			return rows[i].sortKey < rows[j].sortKey
		}
		if rows[i].sortOccasion == rows[j].sortOccasion {
			return rows[i].Occasion < rows[j].Occasion
		}
		return rows[i].sortOccasion < rows[j].sortOccasion
	})
}

// isExactCounterCards reports whether an inventory entry belongs to the exact Counter Cards
// class, which is the only class that should feed the left-hand Data Insights table.
func isExactCounterCards(e *entry) bool {
	category := strings.TrimSpace(e.RawClassDesc)
	if category == "" {
		category = strings.TrimSpace(e.ClassDesc)
	}
	return category == "Counter Cards"
}

// normalizeDataInsightsClassDescription trims the class description and provides a fallback for empty values.
func normalizeDataInsightsClassDescription(e *entry) string {
	category := strings.TrimSpace(e.RawClassDesc)
	if category == "" {
		category = strings.TrimSpace(e.ClassDesc)
	}
	if category == "" {
		return "UNCLASSIFIED"
	}
	return category
}

// setDataInsightsTableWidths applies a shared width pattern starting at the given column.
func setDataInsightsTableWidths(f *excelize.File, sheetName, startCol string, widths []float64) error {
	cols, err := dataInsightsTableColumns(startCol, len(widths))
	if err != nil {
		return err
	}

	for idx, col := range cols {
		if err := f.SetColWidth(sheetName, col, col, widths[idx]); err != nil {
			return fmt.Errorf("failed to set width for column %s: %w", col, err)
		}
	}

	return nil
}

// dataInsightsTableColumns returns table columns starting at startCol.
func dataInsightsTableColumns(startCol string, columnCount int) ([]string, error) {
	startIdx, err := excelize.ColumnNameToNumber(startCol)
	if err != nil {
		return nil, fmt.Errorf("failed to resolve column %s: %w", startCol, err)
	}

	cols := make([]string, columnCount)
	for i := 0; i < len(cols); i++ {
		col, err := excelize.ColumnNumberToName(startIdx + i)
		if err != nil {
			return nil, fmt.Errorf("failed to resolve column %d: %w", startIdx+i, err)
		}
		cols[i] = col
	}

	return cols, nil
}

// normalizeDataInsightsOccasion trims the occasion name and provides a fallback for empty values.
func normalizeDataInsightsOccasion(occ string) string {
	trimmed := strings.TrimSpace(occ)
	if trimmed == "" {
		return "NO OCCASION"
	}
	return trimmed
}

// dataInsightDateInfo returns the display and projection metadata for a data-insight occasion.
func dataInsightDateInfo(section, occasion string, now time.Time) occasionDateInfo {
	if section == "Everyday" {
		return occasionDateInfo{Display: "N/A", SortKey: 999999, TargetMonthsThrough: 12.0}
	}

	// Unknown seasonal occasions fall back to a neutral display and sort after known holidays.
	info, ok := dataInsightDateMap[strings.ToUpper(strings.TrimSpace(occasion))]
	if !ok {
		return occasionDateInfo{Display: "N/A", SortKey: 999999, TargetMonthsThrough: 12.0}
	}

	if isValentinesOccasion(occasion) {
		// Valentine's Day is the one seasonal occasion that is sold in two merchandising waves:
		// an early-year run through Feb 14, then a late-year run from Nov 16 through Dec 31.
		// We show that split directly in the sheet and keep it marked incomplete until year end,
		// because December inventory is still part of the same selling season.
		info.Display = "Jan 1 - Feb 14, Nov 16 - Dec 31"
		info.Complete = !now.Before(time.Date(now.Year()+1, time.January, 1, 0, 0, 0, 0, now.Location()))
		info.TargetMonthsThrough = monthsThroughForDate(now.Year(), info.Month, info.Day, now.Location())
		return info
	}

	if section == "Winter" {
		// Winter holidays don't really start selling until July 1, so treat that as the
		// beginning of the season when calculating the projection window.
		info.TargetMonthsThrough = monthsThroughSinceDate(now.Year(), time.July, 1, info.Month, info.Day, now.Location())
	} else {
		info.TargetMonthsThrough = monthsThroughForDate(now.Year(), info.Month, info.Day, now.Location())
	}

	// Treat the occasion as complete for the whole event day, not just after midnight.
	eventDate := time.Date(now.Year(), info.Month, info.Day, 23, 59, 59, 0, now.Location())
	info.Complete = !now.Before(eventDate)
	return info
}

// isValentinesOccasion centralizes the occasion matching so both spelling variants
// take the same split-window projection path.
func isValentinesOccasion(occasion string) bool {
	switch strings.ToUpper(strings.TrimSpace(occasion)) {
	case "VALENTINE'S DAY", "VALENTINES DAY":
		return true
	default:
		return false
	}
}

// valentinesProjectionWindow returns the active and total selling-day counts for the
// split Valentine's season so projections can scale against the portion of the year
// that is actually on sale instead of assuming the occasion sells all year.
func valentinesProjectionWindow(now time.Time) (currentSellingDays float64, totalSellingDays float64, complete bool) {
	year := now.Year()
	loc := now.Location()

	firstStart := time.Date(year, time.January, 1, 0, 0, 0, 0, loc)
	firstEnd := time.Date(year, time.February, 14, 23, 59, 59, 0, loc)
	secondStart := time.Date(year, time.November, 16, 0, 0, 0, 0, loc)
	yearEnd := time.Date(year+1, time.January, 1, 0, 0, 0, 0, loc)

	firstWindowDays := float64(firstEnd.YearDay() - firstStart.YearDay() + 1)
	secondWindowDays := float64(time.Date(year, time.December, 31, 0, 0, 0, 0, loc).YearDay() - secondStart.YearDay() + 1)
	totalSellingDays = firstWindowDays + secondWindowDays
	complete = !now.Before(yearEnd)

	switch {
	case now.Before(firstStart):
		currentSellingDays = 1
	case !now.After(firstEnd):
		currentSellingDays = float64(now.YearDay() - firstStart.YearDay() + 1)
	case now.Before(secondStart):
		currentSellingDays = firstWindowDays
	default:
		currentSellingDays = firstWindowDays + float64(now.YearDay()-secondStart.YearDay()+1)
	}

	if currentSellingDays <= 0 {
		currentSellingDays = 1
	}
	return currentSellingDays, totalSellingDays, complete
}

// monthsThroughSinceDate returns the month progress between two calendar dates.
func monthsThroughSinceDate(year int, startMonth time.Month, startDay int, endMonth time.Month, endDay int, loc *time.Location) float64 {
	monthsThrough := monthsThroughForDate(year, endMonth, endDay, loc) - monthsThroughForDate(year, startMonth, startDay, loc)
	if monthsThrough <= 0 {
		monthsThrough = 1
	}
	return monthsThrough
}

// formatSeasonStatusYoY formats the season status text with a YoY comparison.
func formatSeasonStatusYoY(projectedSales float64, pySales float64, started bool, complete bool) string {
	if !started {
		return fmt.Sprintf("NOT STARTED: %s YoY", formatYoYFromProjectedSales(projectedSales, pySales))
	}
	if complete {
		return fmt.Sprintf("COMPLETE: %s YoY", formatYoYFromProjectedSales(projectedSales, pySales))
	}
	return fmt.Sprintf("IN PROGRESS: %s YoY", formatYoYFromProjectedSales(projectedSales, pySales))
}

// currentMonthsThrough returns the current year-to-date month progress as a fractional month count.
func currentMonthsThrough(now time.Time) float64 {
	// Use the day-of-month as a fraction so projections move smoothly within the current month.
	year := now.Year()
	month := now.Month()
	daysInMonth := time.Date(year, month+1, 0, 0, 0, 0, 0, now.Location()).Day()
	monthsThrough := float64(int(month)-1) + float64(now.Day())/float64(daysInMonth)
	if monthsThrough <= 0 {
		// Clamp to 1 so early-month dates and unexpected time values still produce usable projections.
		monthsThrough = 1
	}
	return monthsThrough
}

// monthsThroughForDate returns the month progress for a specific calendar date.
func monthsThroughForDate(year int, month time.Month, day int, loc *time.Location) float64 {
	daysInMonth := time.Date(year, month+1, 0, 0, 0, 0, 0, loc).Day()
	monthsThrough := float64(int(month)-1) + float64(day)/float64(daysInMonth)
	if monthsThrough <= 0 {
		// Clamp to 1 so the projection math never divides by zero.
		monthsThrough = 1
	}
	return monthsThrough
}

// formatYoYFromProjectedSales calculates and formats the YoY percentage from projected and prior-year sales.
func formatYoYFromProjectedSales(projectedSales float64, pySales float64) string {
	// Avoid divide-by-zero when there is no prior-year baseline to compare against.
	if pySales == 0 {
		return "N/A"
	}
	pct := math.Round(((projectedSales - pySales) / pySales) * 100)
	return fmt.Sprintf("%+.0f%%", pct)
}
