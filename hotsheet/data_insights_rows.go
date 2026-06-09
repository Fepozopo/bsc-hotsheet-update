package hotsheet

import (
	"sort"
	"strings"
	"time"
)

// dataInsightsRow represents one output row in the Data Insights sheet.
//
// YoYDisplay stores the exact text written into the rightmost column. For seasonal rows that
// includes status wording such as "IN PROGRESS:" or "COMPLETE:", while everyday rows only
// show the projected year-over-year percentage.
type dataInsightsRow struct {
	Class               string
	Occasion            string
	Date                string
	DollarSoldYTD       float64
	DollarSoldPY        float64
	ProjectedDollar     float64
	YoYDisplay          string
	occasionDateSortKey int
	occasionSortLabel   string
}

// dataInsightsGroup aggregates sales for a single normalized occasion within a section.
type dataInsightsGroup struct {
	Section             string
	Class               string
	Occasion            string
	Date                string
	occasionDateSortKey int
	complete            bool
	TargetMonthsThrough float64
	DollarSoldYTD       float64
	DollarSoldPY        float64
}

// buildDataInsightsRows groups Counter Cards into the seasonal sections used by the
// Data Insights sheet, preserving the existing holiday/date/projection rules.
func buildDataInsightsRows(entries []*inventoryEntry, currentMonthsThrough float64, now time.Time) map[string][]dataInsightsRow {
	groups := make(map[string]*dataInsightsGroup)

	for _, e := range entries {
		// This sheet only tracks exact Counter Cards entries.
		if !isExactCounterCards(e) {
			continue
		}

		// Normalize occasion names so common variants collapse into the same grouped row.
		occasion := normalizeDataInsightsOccasion(e.Occasion)
		section := mapOccasion(occasion)
		dateInfo := dataInsightsDateInfo(section, occasion, now)
		// Use the normalized occasion plus section so variants collapse into one rollup bucket.
		groupKey := section + "|" + strings.ToUpper(occasion)

		group, ok := groups[groupKey]
		if !ok {
			group = &dataInsightsGroup{
				Section:             section,
				Occasion:            occasion,
				Date:                dateInfo.Display,
				occasionDateSortKey: dateInfo.SortKey,
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
		dateInfo, projected, yoyDisplay := projectDataInsightsRow(group.Section, group.Occasion, group.DollarSoldYTD, group.DollarSoldPY, currentMonthsThrough, now)
		row := dataInsightsRow{
			Occasion:            group.Occasion,
			Date:                dateInfo.Display,
			DollarSoldYTD:       group.DollarSoldYTD,
			DollarSoldPY:        group.DollarSoldPY,
			ProjectedDollar:     projected,
			YoYDisplay:          yoyDisplay,
			occasionDateSortKey: group.occasionDateSortKey,
			occasionSortLabel:   strings.ToUpper(group.Occasion),
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

// buildOtherProductsDataInsightsRows groups non-card products by class and occasion, then
// applies the same seasonal bucketing and projection rules used by the card section.
func buildOtherProductsDataInsightsRows(entries []*inventoryEntry, currentMonthsThrough float64, now time.Time) map[string][]dataInsightsRow {
	groups := make(map[string]*dataInsightsGroup)

	for _, e := range entries {
		if isExactCounterCards(e) {
			continue
		}

		classDesc := normalizeDataInsightsClassDescription(e)
		occasion := normalizeDataInsightsOccasion(e.Occasion)
		section := mapOccasion(occasion)
		dateInfo := dataInsightsDateInfo(section, occasion, now)
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
				occasionDateSortKey: dateInfo.SortKey,
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
		dateInfo, projected, yoyDisplay := projectDataInsightsRow(group.Section, group.Occasion, group.DollarSoldYTD, group.DollarSoldPY, currentMonthsThrough, now)
		row := dataInsightsRow{
			Class:               group.Class,
			Occasion:            group.Occasion,
			Date:                dateInfo.Display,
			DollarSoldYTD:       group.DollarSoldYTD,
			DollarSoldPY:        group.DollarSoldPY,
			ProjectedDollar:     projected,
			YoYDisplay:          yoyDisplay,
			occasionDateSortKey: group.occasionDateSortKey,
			occasionSortLabel:   strings.ToUpper(group.Occasion),
		}
		rowsBySection[group.Section] = append(rowsBySection[group.Section], row)
	}

	sortDataInsightsRows(rowsBySection["Spring"], true, true)
	sortDataInsightsRows(rowsBySection["Winter"], true, true)
	sortDataInsightsRows(rowsBySection["Everyday"], false, true)

	return rowsBySection
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
		if useSortKey && rows[i].occasionDateSortKey != rows[j].occasionDateSortKey {
			return rows[i].occasionDateSortKey < rows[j].occasionDateSortKey
		}
		if rows[i].occasionSortLabel == rows[j].occasionSortLabel {
			return rows[i].Occasion < rows[j].Occasion
		}
		return rows[i].occasionSortLabel < rows[j].occasionSortLabel
	})
}

// isExactCounterCards reports whether an inventory entry belongs to the exact Counter Cards
// class, which is the only class that should feed the left-hand Data Insights table.
func isExactCounterCards(item *inventoryEntry) bool {
	category := strings.TrimSpace(item.RawClassDesc)
	if category == "" {
		category = strings.TrimSpace(item.ClassDesc)
	}
	return category == "Counter Cards"
}

// normalizeDataInsightsClassDescription trims the class description and provides a fallback for empty values.
func normalizeDataInsightsClassDescription(item *inventoryEntry) string {
	category := strings.TrimSpace(item.RawClassDesc)
	if category == "" {
		category = strings.TrimSpace(item.ClassDesc)
	}
	if category == "" {
		return "UNCLASSIFIED"
	}
	return category
}

// normalizeDataInsightsOccasion trims the occasion name and provides a fallback for empty values.
func normalizeDataInsightsOccasion(occ string) string {
	trimmed := strings.TrimSpace(occ)
	if trimmed == "" {
		return "NO OCCASION"
	}
	return trimmed
}
