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
	dataInsightsSheetName          = "Data Insights"
	dataInsightsTitleText          = "Data Insights"
	dataInsightsLeftTableStartCol  = "A"
	dataInsightsRightTableStartCol = "G"
	dataInsightsTitleRow           = 1
	dataInsightsSectionRow         = 3
	dataInsightsTableStartRow      = 5
	dataInsightsTableColumnCount   = 5
)

const (
	dataInsightsColumnOccasion = iota
	dataInsightsColumnDate
	dataInsightsColumnYTD
	dataInsightsColumnPY
	dataInsightsColumnFinal
)

var dataInsightsTableColumnWidths = []float64{26, 31, 14, 14, 34}

// dataInsightsCell constructs an Excel cell name (e.g. "A5") from a column letter and row number.
func dataInsightsCell(col string, row int) string {
	return fmt.Sprintf("%s%d", col, row)
}

// dataInsightsTableEndCol calculates the last column of a Data Insights table based on the shared column count.
func dataInsightsTableEndCol(startCol string) (string, error) {
	cols, err := dataInsightsTableColumns(startCol)
	if err != nil {
		return "", err
	}

	return cols[dataInsightsTableColumnCount-1], nil
}

// dataInsightsRow represents one output row in the Data Insights sheet.
type dataInsightsRow struct {
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
	"VALENTINE'S DAY":   {Display: "February 14", Month: time.February, Day: 14, SortKey: 214},
	"VALENTINES DAY":    {Display: "February 14", Month: time.February, Day: 14, SortKey: 214},
	"ST PATRICKS DAY":   {Display: "March 17", Month: time.March, Day: 17, SortKey: 317},
	"ST. PATRICK'S DAY": {Display: "March 17", Month: time.March, Day: 17, SortKey: 317},
	"EASTER":            {Display: "April 20", Month: time.April, Day: 20, SortKey: 420},
	"MOTHER'S DAY":      {Display: "May 11", Month: time.May, Day: 11, SortKey: 511},
	"MOTHERS DAY":       {Display: "May 11", Month: time.May, Day: 11, SortKey: 511},
	"GRADUATION":        {Display: "June", Month: time.June, Day: 30, SortKey: 630},
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
	otherRows := buildOtherProductDataInsightsRows(entries, currentMonthsThrough)

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

	if err := setDataInsightsTableWidths(f, sheetName, dataInsightsLeftTableStartCol); err != nil {
		return err
	}
	if err := setDataInsightsTableWidths(f, sheetName, dataInsightsRightTableStartCol); err != nil {
		return err
	}

	leftCols, err := dataInsightsTableColumns(dataInsightsLeftTableStartCol)
	if err != nil {
		return err
	}
	leftEndCol := leftCols[dataInsightsTableColumnCount-1]
	titleEndCol, err := dataInsightsTableEndCol(dataInsightsRightTableStartCol)
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

	// Leave a blank row between the subtitles and the tables for readability.
	rowNum := dataInsightsTableStartRow
	// Spring and Winter use completion-aware status text. Everyday uses a straight projected YoY value.
	sections := []struct {
		name    string
		headers []string
	}{
		{name: "Spring", headers: []string{"Occasion", "Date", "YTD Sales", "PY Sales", "Status / Projected YoY"}},
		{name: "Winter", headers: []string{"Occasion", "Date", "YTD Sales", "PY Sales", "Status / Projected YoY"}},
		{name: "Everyday", headers: []string{"Occasion", "Date", "YTD Sales", "PY Sales", "Projected YoY"}},
	}

	for idx, section := range sections {
		if idx > 0 {
			rowNum++
		}
		sectionTitleCell := dataInsightsCell(dataInsightsLeftTableStartCol, rowNum)
		sectionTitleEnd := dataInsightsCell(leftEndCol, rowNum)
		if err := f.SetCellValue(sheetName, sectionTitleCell, section.name); err != nil {
			return fmt.Errorf("failed to set section title %s: %w", section.name, err)
		}
		if err := f.MergeCell(sheetName, sectionTitleCell, sectionTitleEnd); err != nil {
			return fmt.Errorf("failed to merge section title %s: %w", section.name, err)
		}
		if err := f.SetCellStyle(sheetName, sectionTitleCell, sectionTitleEnd, sectionStyle); err != nil {
			return fmt.Errorf("failed to style section title %s: %w", section.name, err)
		}
		rowNum++

		headers := section.headers
		for colIdx, header := range headers {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowNum)
			if err := f.SetCellValue(sheetName, cell, header); err != nil {
				return fmt.Errorf("failed to set header %s: %w", header, err)
			}
		}
		if err := f.SetCellStyle(sheetName, dataInsightsCell(leftCols[dataInsightsColumnOccasion], rowNum), dataInsightsCell(leftEndCol, rowNum), headerStyle); err != nil {
			return fmt.Errorf("failed to style header row for %s: %w", section.name, err)
		}
		rowNum++

		sectionRows := rowsBySection[section.name]
		totalYTD := 0.0
		totalPY := 0.0
		totalProjectedSales := 0.0
		// A section is considered complete only when none of its rows are still in progress.
		sectionComplete := true
		for _, row := range sectionRows {
			values := []interface{}{row.Occasion, row.Date, row.DollarSoldYTD, row.DollarSoldPY, row.Final}
			for colIdx, value := range values {
				cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowNum)
				if err := f.SetCellValue(sheetName, cell, value); err != nil {
					return fmt.Errorf("failed to write %s row cell %s: %w", section.name, cell, err)
				}
			}
			if err := f.SetCellStyle(sheetName, dataInsightsCell(leftCols[dataInsightsColumnOccasion], rowNum), dataInsightsCell(leftEndCol, rowNum), dataStyle); err != nil {
				return fmt.Errorf("failed to style %s data row: %w", section.name, err)
			}
			if err := f.SetCellStyle(sheetName, dataInsightsCell(leftCols[dataInsightsColumnYTD], rowNum), dataInsightsCell(leftCols[dataInsightsColumnPY], rowNum), currencyDataStyle); err != nil {
				return fmt.Errorf("failed to style %s currency cells: %w", section.name, err)
			}
			totalYTD += row.DollarSoldYTD
			totalPY += row.DollarSoldPY
			totalProjectedSales += row.ProjectedDollar
			// If any row is still in progress, the section total should use the same wording.
			if strings.HasPrefix(row.Final, "IN PROGRESS:") {
				sectionComplete = false
			}
			rowNum++
		}

		totalRowValues := []interface{}{"Total", "", totalYTD, totalPY, ""}
		// Totals preserve the same YoY wording style used by the section's detailed rows.
		if section.name == "Spring" || section.name == "Winter" {
			totalRowValues[dataInsightsColumnFinal] = formatSeasonStatusYoY(totalProjectedSales, totalPY, sectionComplete)
		}
		if section.name == "Everyday" {
			totalRowValues[dataInsightsColumnFinal] = formatYoYFromProjectedSales(totalProjectedSales, totalPY)
		}
		for colIdx, value := range totalRowValues {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowNum)
			if err := f.SetCellValue(sheetName, cell, value); err != nil {
				return fmt.Errorf("failed to write %s total row cell %s: %w", section.name, cell, err)
			}
		}
		if err := f.SetCellStyle(sheetName, dataInsightsCell(leftCols[dataInsightsColumnOccasion], rowNum), dataInsightsCell(leftEndCol, rowNum), totalStyle); err != nil {
			return fmt.Errorf("failed to style total row for %s: %w", section.name, err)
		}
		if err := f.SetCellStyle(sheetName, dataInsightsCell(leftCols[dataInsightsColumnYTD], rowNum), dataInsightsCell(leftCols[dataInsightsColumnPY], rowNum), currencyTotalStyle); err != nil {
			return fmt.Errorf("failed to style %s total currency cells: %w", section.name, err)
		}
		rowNum++
	}

	if err := writeOtherProductsTable(f, sheetName, dataInsightsRightTableStartCol, dataInsightsTableStartRow, otherRows, headerStyle, dataStyle, currencyDataStyle, totalStyle, currencyTotalStyle); err != nil {
		return err
	}

	return nil
}

// buildDataInsightsRows groups qualifying entries into the rows used by the Data Insights sheet.
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

	// Pre-create all sections so the sheet still renders empty categories consistently.
	rowsBySection := map[string][]dataInsightsRow{
		"Spring":   {},
		"Winter":   {},
		"Everyday": {},
	}

	for _, group := range groups {
		row := dataInsightsRow{
			Occasion:      group.Occasion,
			Date:          group.Date,
			DollarSoldYTD: group.DollarSoldYTD,
			DollarSoldPY:  group.DollarSoldPY,
			sortKey:       group.sortKey,
			sortOccasion:  strings.ToUpper(group.Occasion),
		}
		if group.Section == "Spring" || group.Section == "Winter" {
			// Seasonal items are projected only up to their typical selling window.
			if isValentinesOccasion(group.Occasion) {
				currentSellingDays, totalSellingDays, complete := valentinesProjectionWindow(now)
				row.ProjectedDollar = group.DollarSoldYTD * (totalSellingDays / currentSellingDays)
				row.Final = formatSeasonStatusYoY(row.ProjectedDollar, group.DollarSoldPY, complete)
			} else if group.complete {
				row.ProjectedDollar = group.DollarSoldYTD
				row.Final = formatSeasonStatusYoY(row.ProjectedDollar, group.DollarSoldPY, group.complete)
			} else {
				row.ProjectedDollar = group.DollarSoldYTD * (group.TargetMonthsThrough / currentMonthsThrough)
				row.Final = formatSeasonStatusYoY(row.ProjectedDollar, group.DollarSoldPY, group.complete)
			}
		} else {
			// Everyday items do not have a calendar cutover, so they project across the full year.
			row.Date = "N/A"
			row.ProjectedDollar = group.DollarSoldYTD * (12.0 / currentMonthsThrough)
			row.Final = formatYoYFromProjectedSales(row.ProjectedDollar, group.DollarSoldPY)
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
	sort.Slice(rowsBySection["Spring"], func(i, j int) bool {
		if rowsBySection["Spring"][i].sortKey == rowsBySection["Spring"][j].sortKey {
			return rowsBySection["Spring"][i].sortOccasion < rowsBySection["Spring"][j].sortOccasion
		}
		return rowsBySection["Spring"][i].sortKey < rowsBySection["Spring"][j].sortKey
	})
	sort.Slice(rowsBySection["Winter"], func(i, j int) bool {
		if rowsBySection["Winter"][i].sortKey == rowsBySection["Winter"][j].sortKey {
			return rowsBySection["Winter"][i].sortOccasion < rowsBySection["Winter"][j].sortOccasion
		}
		return rowsBySection["Winter"][i].sortKey < rowsBySection["Winter"][j].sortKey
	})
	sort.Slice(rowsBySection["Everyday"], func(i, j int) bool {
		return rowsBySection["Everyday"][i].sortOccasion < rowsBySection["Everyday"][j].sortOccasion
	})

	return rowsBySection
}

// dataInsightsClassRow represents one output row in the Other Products table.
type dataInsightsClassRow struct {
	Class           string
	Date            string
	DollarSoldYTD   float64
	DollarSoldPY    float64
	ProjectedDollar float64
	Final           string
}

// dataInsightsClassGroup aggregates sales for a single non-counter-card class description.
type dataInsightsClassGroup struct {
	Class         string
	DollarSoldYTD float64
	DollarSoldPY  float64
}

// buildOtherProductDataInsightsRows groups all non-counter-card entries by class description.
func buildOtherProductDataInsightsRows(entries []*entry, currentMonthsThrough float64) []dataInsightsClassRow {
	groups := make(map[string]*dataInsightsClassGroup)

	for _, e := range entries {
		if isExactCounterCards(e) {
			continue
		}

		classDesc := normalizeDataInsightsClassDescription(e)
		groupKey := strings.ToUpper(classDesc)

		group, ok := groups[groupKey]
		if !ok {
			group = &dataInsightsClassGroup{Class: classDesc}
			groups[groupKey] = group
		}

		group.DollarSoldYTD += e.DollarSoldYTD
		group.DollarSoldPY += e.DollarSoldPY
	}

	rows := make([]dataInsightsClassRow, 0, len(groups))
	for _, group := range groups {
		projected := group.DollarSoldYTD * (12.0 / currentMonthsThrough)
		rows = append(rows, dataInsightsClassRow{
			Class:           group.Class,
			Date:            "N/A",
			DollarSoldYTD:   group.DollarSoldYTD,
			DollarSoldPY:    group.DollarSoldPY,
			ProjectedDollar: projected,
			Final:           formatYoYFromProjectedSales(projected, group.DollarSoldPY),
		})
	}

	sort.Slice(rows, func(i, j int) bool {
		return strings.ToUpper(rows[i].Class) < strings.ToUpper(rows[j].Class)
	})

	return rows
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

// setDataInsightsTableWidths applies the shared 5-column width pattern starting at the given column.
func setDataInsightsTableWidths(f *excelize.File, sheetName, startCol string) error {
	cols, err := dataInsightsTableColumns(startCol)
	if err != nil {
		return err
	}

	if len(dataInsightsTableColumnWidths) != dataInsightsTableColumnCount {
		return fmt.Errorf("invalid data insights table width count: got %d, want %d", len(dataInsightsTableColumnWidths), dataInsightsTableColumnCount)
	}

	for idx, col := range cols {
		if err := f.SetColWidth(sheetName, col, col, dataInsightsTableColumnWidths[idx]); err != nil {
			return fmt.Errorf("failed to set width for column %s: %w", col, err)
		}
	}

	return nil
}

// dataInsightsTableColumns returns the shared table columns starting at startCol.
func dataInsightsTableColumns(startCol string) ([]string, error) {
	startIdx, err := excelize.ColumnNameToNumber(startCol)
	if err != nil {
		return nil, fmt.Errorf("failed to resolve column %s: %w", startCol, err)
	}

	cols := make([]string, dataInsightsTableColumnCount)
	for i := 0; i < len(cols); i++ {
		col, err := excelize.ColumnNumberToName(startIdx + i)
		if err != nil {
			return nil, fmt.Errorf("failed to resolve column %d: %w", startIdx+i, err)
		}
		cols[i] = col
	}

	return cols, nil
}

// writeOtherProductsTable writes the right-side Other Products table.
func writeOtherProductsTable(f *excelize.File, sheetName, startCol string, startRow int, rows []dataInsightsClassRow, headerStyle, dataStyle, currencyDataStyle, totalStyle, currencyTotalStyle int) error {
	cols, err := dataInsightsTableColumns(startCol)
	if err != nil {
		return err
	}

	headers := []string{"Class", "Date", "YTD Sales", "PY Sales", "Projected YoY"}
	for idx, header := range headers {
		cell := dataInsightsCell(cols[idx], startRow)
		if err := f.SetCellValue(sheetName, cell, header); err != nil {
			return fmt.Errorf("failed to set other products header %s: %w", header, err)
		}
	}
	if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[dataInsightsColumnOccasion], startRow), dataInsightsCell(cols[dataInsightsColumnFinal], startRow), headerStyle); err != nil {
		return fmt.Errorf("failed to style other products header row: %w", err)
	}

	rowNum := startRow + 1
	totalYTD := 0.0
	totalPY := 0.0
	totalProjectedSales := 0.0
	for _, row := range rows {
		values := []interface{}{row.Class, row.Date, row.DollarSoldYTD, row.DollarSoldPY, row.Final}
		for idx, value := range values {
			cell := dataInsightsCell(cols[idx], rowNum)
			if err := f.SetCellValue(sheetName, cell, value); err != nil {
				return fmt.Errorf("failed to write other products row cell %s: %w", cell, err)
			}
		}
		if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[dataInsightsColumnOccasion], rowNum), dataInsightsCell(cols[dataInsightsColumnFinal], rowNum), dataStyle); err != nil {
			return fmt.Errorf("failed to style other products data row: %w", err)
		}
		if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[dataInsightsColumnYTD], rowNum), dataInsightsCell(cols[dataInsightsColumnPY], rowNum), currencyDataStyle); err != nil {
			return fmt.Errorf("failed to style other products currency cells: %w", err)
		}

		totalYTD += row.DollarSoldYTD
		totalPY += row.DollarSoldPY
		totalProjectedSales += row.ProjectedDollar
		rowNum++
	}

	totalRowValues := []interface{}{"Total", "", totalYTD, totalPY, formatYoYFromProjectedSales(totalProjectedSales, totalPY)}
	for idx, value := range totalRowValues {
		cell := dataInsightsCell(cols[idx], rowNum)
		if err := f.SetCellValue(sheetName, cell, value); err != nil {
			return fmt.Errorf("failed to write other products total row cell %s: %w", cell, err)
		}
	}
	if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[dataInsightsColumnOccasion], rowNum), dataInsightsCell(cols[dataInsightsColumnFinal], rowNum), totalStyle); err != nil {
		return fmt.Errorf("failed to style other products total row: %w", err)
	}
	if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[dataInsightsColumnYTD], rowNum), dataInsightsCell(cols[dataInsightsColumnPY], rowNum), currencyTotalStyle); err != nil {
		return fmt.Errorf("failed to style other products total currency cells: %w", err)
	}

	return nil
}

// isExactCounterCards reports whether the entry belongs to the exact "Counter Cards" class.
func isExactCounterCards(e *entry) bool {
	category := strings.TrimSpace(e.RawClassDesc)
	if category == "" {
		category = strings.TrimSpace(e.ClassDesc)
	}
	return category == "Counter Cards"
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
		info.Display = "Jan 1 - Feb 14, Nov 16 - Dec 31"
		info.Complete = !now.Before(time.Date(now.Year()+1, time.January, 1, 0, 0, 0, 0, now.Location()))
		info.TargetMonthsThrough = monthsThroughForDate(now.Year(), info.Month, info.Day, now.Location())
		return info
	}

	// Treat the occasion as complete for the whole event day, not just after midnight.
	eventDate := time.Date(now.Year(), info.Month, info.Day, 23, 59, 59, 0, now.Location())
	info.Complete = !now.Before(eventDate)
	info.TargetMonthsThrough = monthsThroughForDate(now.Year(), info.Month, info.Day, now.Location())
	return info
}

func isValentinesOccasion(occasion string) bool {
	switch strings.ToUpper(strings.TrimSpace(occasion)) {
	case "VALENTINE'S DAY", "VALENTINES DAY":
		return true
	default:
		return false
	}
}

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

// formatSeasonStatusYoY formats the season status text with a YoY comparison.
func formatSeasonStatusYoY(projectedSales float64, pySales float64, complete bool) string {
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
