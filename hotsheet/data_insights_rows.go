package hotsheet

import (
	"sort"
	"strings"
	"time"
)

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
		dateInfo := dataInsightsDateInfo(section, occasion, now)
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

// normalizeDataInsightsOccasion trims the occasion name and provides a fallback for empty values.
func normalizeDataInsightsOccasion(occ string) string {
	trimmed := strings.TrimSpace(occ)
	if trimmed == "" {
		return "NO OCCASION"
	}
	return trimmed
}
