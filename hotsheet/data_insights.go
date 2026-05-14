package hotsheet

import (
	"fmt"
	"math"
	"sort"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type dataInsightsRow struct {
	Occasion     string
	Date         string
	YTDSales     int
	PYSales      int
	Final        string
	sortKey      int
	sortOccasion string
}

type dataInsightsGroup struct {
	Section             string
	Occasion            string
	Date                string
	sortKey             int
	complete            bool
	TargetMonthsThrough float64
	YTDSales            int
	PYSales             int
}

type occasionDateInfo struct {
	Display             string
	Month               time.Month
	Day                 int
	SortKey             int
	Complete            bool
	TargetMonthsThrough float64
}

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
	"HOLIDAY":           {Display: "December 15", Month: time.December, Day: 15, SortKey: 1215},
	"CHRISTMAS":         {Display: "December 25", Month: time.December, Day: 25, SortKey: 1225},
}

func writeDataInsightsSheet(f *excelize.File, entries []*entry) error {
	currentMonthsThrough := currentMonthsThrough()
	rowsBySection := buildDataInsightsRows(entries, currentMonthsThrough)

	sheetName := "Data Insights"
	if _, err := f.NewSheet(sheetName); err != nil {
		return fmt.Errorf("failed to create %s sheet: %w", sheetName, err)
	}

	titleStyle, err := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 14},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	if err != nil {
		return fmt.Errorf("failed to create title style: %w", err)
	}

	sectionStyle, err := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#D9EAF7"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create section style: %w", err)
	}

	headerStyle, err := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#E6E6FA"}, Pattern: 1},
	})
	if err != nil {
		return fmt.Errorf("failed to create header style: %w", err)
	}

	dataStyle, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create data style: %w", err)
	}

	totalStyle, err := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#F2F2F2"}, Pattern: 1},
	})
	if err != nil {
		return fmt.Errorf("failed to create total style: %w", err)
	}

	columnWidths := map[string]float64{
		"A": 26,
		"B": 16,
		"C": 14,
		"D": 14,
		"E": 34,
	}
	for col, width := range columnWidths {
		if err := f.SetColWidth(sheetName, col, col, width); err != nil {
			return fmt.Errorf("failed to set width for column %s: %w", col, err)
		}
	}

	if err := f.SetCellValue(sheetName, "A1", "Data Insights"); err != nil {
		return fmt.Errorf("failed to set sheet title: %w", err)
	}
	if err := f.MergeCell(sheetName, "A1", "E1"); err != nil {
		return fmt.Errorf("failed to merge title cells: %w", err)
	}
	if err := f.SetCellStyle(sheetName, "A1", "E1", titleStyle); err != nil {
		return fmt.Errorf("failed to style title row: %w", err)
	}

	rowNum := 3
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
		sectionTitleCell := fmt.Sprintf("A%d", rowNum)
		sectionTitleEnd := fmt.Sprintf("E%d", rowNum)
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
		if err := f.SetCellStyle(sheetName, fmt.Sprintf("A%d", rowNum), fmt.Sprintf("E%d", rowNum), headerStyle); err != nil {
			return fmt.Errorf("failed to style header row for %s: %w", section.name, err)
		}
		rowNum++

		sectionRows := rowsBySection[section.name]
		totalYTD := 0
		totalPY := 0
		for _, row := range sectionRows {
			values := []interface{}{row.Occasion, row.Date, row.YTDSales, row.PYSales, row.Final}
			for colIdx, value := range values {
				cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowNum)
				if err := f.SetCellValue(sheetName, cell, value); err != nil {
					return fmt.Errorf("failed to write %s row cell %s: %w", section.name, cell, err)
				}
			}
			if err := f.SetCellStyle(sheetName, fmt.Sprintf("A%d", rowNum), fmt.Sprintf("E%d", rowNum), dataStyle); err != nil {
				return fmt.Errorf("failed to style %s data row: %w", section.name, err)
			}
			totalYTD += row.YTDSales
			totalPY += row.PYSales
			rowNum++
		}

		totalRowValues := []interface{}{"Total", "", totalYTD, totalPY, ""}
		if section.name == "Everyday" {
			totalRowValues[4] = formatProjectedYoY(totalYTD, totalPY, currentMonthsThrough, 12.0)
		}
		for colIdx, value := range totalRowValues {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowNum)
			if err := f.SetCellValue(sheetName, cell, value); err != nil {
				return fmt.Errorf("failed to write %s total row cell %s: %w", section.name, cell, err)
			}
		}
		if err := f.SetCellStyle(sheetName, fmt.Sprintf("A%d", rowNum), fmt.Sprintf("E%d", rowNum), totalStyle); err != nil {
			return fmt.Errorf("failed to style total row for %s: %w", section.name, err)
		}
		rowNum++
	}

	return nil
}

func buildDataInsightsRows(entries []*entry, currentMonthsThrough float64) map[string][]dataInsightsRow {
	groups := make(map[string]*dataInsightsGroup)

	for _, e := range entries {
		if !isExactCounterCards(e) {
			continue
		}

		occasion := normalizeDataInsightsOccasion(e.Occasion)
		section := mapOccasion(occasion)
		dateInfo := dataInsightDateInfo(section, occasion)
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

		group.YTDSales += e.YTDSold + max(e.YTDIssued, 0)
		group.PYSales += e.SoldPY + max(e.IssuedPY, 0)
	}

	rowsBySection := map[string][]dataInsightsRow{
		"Spring":   {},
		"Winter":   {},
		"Everyday": {},
	}

	for _, group := range groups {
		row := dataInsightsRow{
			Occasion:     group.Occasion,
			Date:         group.Date,
			YTDSales:     group.YTDSales,
			PYSales:      group.PYSales,
			sortKey:      group.sortKey,
			sortOccasion: strings.ToUpper(group.Occasion),
		}
		if group.Section == "Spring" || group.Section == "Winter" {
			row.Final = formatSeasonStatusYoY(group.YTDSales, group.PYSales, currentMonthsThrough, group.TargetMonthsThrough, group.complete)
		} else {
			row.Date = "N/A"
			row.Final = formatProjectedYoY(group.YTDSales, group.PYSales, currentMonthsThrough, 12.0)
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

func isExactCounterCards(e *entry) bool {
	category := strings.TrimSpace(e.RawClassDesc)
	if category == "" {
		category = strings.TrimSpace(e.ClassDesc)
	}
	return category == "Counter Cards"
}

func normalizeDataInsightsOccasion(occ string) string {
	trimmed := strings.TrimSpace(occ)
	if trimmed == "" {
		return "NO OCCASION"
	}
	return trimmed
}

func dataInsightDateInfo(section, occasion string) occasionDateInfo {
	if section == "Everyday" {
		return occasionDateInfo{Display: "N/A", SortKey: 999999, TargetMonthsThrough: 12.0}
	}

	info, ok := dataInsightDateMap[strings.ToUpper(strings.TrimSpace(occasion))]
	if !ok {
		return occasionDateInfo{Display: "N/A", SortKey: 999999, TargetMonthsThrough: 12.0}
	}

	now := time.Now()
	eventDate := time.Date(now.Year(), info.Month, info.Day, 23, 59, 59, 0, now.Location())
	info.Complete = !now.Before(eventDate)
	info.TargetMonthsThrough = monthsThroughForDate(now.Year(), info.Month, info.Day, now.Location())
	return info
}

func formatSeasonStatusYoY(ytdSales, pySales int, currentMonthsThrough, targetMonthsThrough float64, complete bool) string {
	if complete {
		return fmt.Sprintf("COMPLETE: %s YoY", formatYoYPercent(ytdSales, pySales))
	}
	return fmt.Sprintf("IN PROGRESS: %s YoY", formatProjectedYoYPercent(ytdSales, pySales, currentMonthsThrough, targetMonthsThrough))
}

func currentMonthsThrough() float64 {
	now := time.Now()
	year := now.Year()
	month := now.Month()
	daysInMonth := time.Date(year, month+1, 0, 0, 0, 0, 0, now.Location()).Day()
	monthsThrough := float64(int(month)-1) + float64(now.Day())/float64(daysInMonth)
	if monthsThrough <= 0 {
		monthsThrough = 1
	}
	return monthsThrough
}

func monthsThroughForDate(year int, month time.Month, day int, loc *time.Location) float64 {
	daysInMonth := time.Date(year, month+1, 0, 0, 0, 0, 0, loc).Day()
	monthsThrough := float64(int(month)-1) + float64(day)/float64(daysInMonth)
	if monthsThrough <= 0 {
		monthsThrough = 1
	}
	return monthsThrough
}

func formatProjectedYoY(ytdSales, pySales int, currentMonthsThrough, targetMonthsThrough float64) string {
	return formatProjectedYoYPercent(ytdSales, pySales, currentMonthsThrough, targetMonthsThrough)
}

func formatProjectedYoYPercent(ytdSales, pySales int, currentMonthsThrough, targetMonthsThrough float64) string {
	if pySales == 0 || currentMonthsThrough <= 0 || targetMonthsThrough <= 0 {
		return "N/A"
	}
	projectedSales := float64(ytdSales) * (targetMonthsThrough / currentMonthsThrough)
	pct := math.Round(((projectedSales - float64(pySales)) / float64(pySales)) * 100)
	return fmt.Sprintf("%+.0f%%", pct)
}

func formatYoYPercent(ytdSales, pySales int) string {
	if pySales == 0 {
		return "N/A"
	}
	pct := math.Round((float64(ytdSales-pySales) / float64(pySales)) * 100)
	return fmt.Sprintf("%+.0f%%", pct)
}
