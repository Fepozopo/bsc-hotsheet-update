package hotsheet

import (
	"sort"
	"strings"
	"time"
)

// dataInsightsRow holds one combined occasion's inventory totals and history-based
// outlook. DollarSoldYTD and DollarSoldPY always retain inventory-report values;
// DollarSoldPYToDate and ProjectedDollar use optional monthly dollar history.
// HasProjection distinguishes a real zero estimate from unavailable history.
// Class is empty for the Counter Cards table and names each Other Products table.
type dataInsightsRow struct {
	DollarSoldYTD      float64
	DollarSoldPY       float64
	DollarSoldPYToDate float64
	ProjectedDollar    float64
	Class              string
	Occasion           string
	Date               string
	HasProjection      bool
	LimitedHistory     bool
}

// dataInsightsGroup accumulates a combined row and its observed yearly dollars.
// A year exists only when a member SKU supplies a Dollars Sold record for it.
type dataInsightsGroup struct {
	row     dataInsightsRow
	history map[int][12]float64
}

// dataInsightsGroupKey keeps the class/occasion aggregation boundary explicit;
// separate fields avoid collisions when source labels contain punctuation.
type dataInsightsGroupKey struct {
	class    string
	occasion string
}

// buildDataInsightsRows returns Counter Cards grouped by occasion and Other
// Products grouped by class then occasion. Labels are trimmed and grouped
// case-insensitively; tables and rows are alphabetically ordered by the writer.
// Nil entries are ignored. hasHistory enables projections anchored exclusively
// to runDate; without history, inventory totals remain but projections do not.
func buildDataInsightsRows(entries []*inventoryEntry, hasHistory bool, runDate time.Time) ([]dataInsightsRow, map[string][]dataInsightsRow) {
	groups := make(map[dataInsightsGroupKey]*dataInsightsGroup)
	classLabels := make(map[string]string)
	for _, entry := range entries {
		if entry == nil {
			continue
		}
		class := ""
		if !isExactCounterCards(entry) {
			class = normalizeDataInsightsClassDescription(entry)
			key := strings.ToUpper(class)
			if label, exists := classLabels[key]; exists {
				class = label
			} else {
				classLabels[key] = class
			}
		}
		occasion := normalizeDataInsightsOccasion(entry.Occasion)
		key := dataInsightsGroupKey{class: strings.ToUpper(class), occasion: strings.ToUpper(occasion)}
		group := groups[key]
		if group == nil {
			group = &dataInsightsGroup{
				row:     dataInsightsRow{Class: class, Occasion: occasion, Date: dataInsightsHolidayLabel(occasion)},
				history: make(map[int][12]float64),
			}
			groups[key] = group
		}
		group.row.DollarSoldYTD += entry.DollarSoldYTD
		group.row.DollarSoldPY += entry.DollarSoldPY
		if hasHistory {
			for _, record := range entry.SalesRecords {
				if record.Metric != "Dollars Sold" {
					continue
				}
				dollars := group.history[record.Year]
				for month, value := range record.Periods {
					dollars[month] += value
				}
				group.history[record.Year] = dollars
			}
		}
	}

	var cards []dataInsightsRow
	other := make(map[string][]dataInsightsRow)
	for _, group := range groups {
		if hasHistory && !runDate.IsZero() {
			projectDataInsightsRow(&group.row, group.history, runDate)
		}
		if group.row.Class == "" {
			cards = append(cards, group.row)
		} else {
			other[group.row.Class] = append(other[group.row.Class], group.row)
		}
	}
	sortDataInsightsRows(cards)
	for _, rows := range other {
		sortDataInsightsRows(rows)
	}
	return cards, other
}

// sortedOtherProductsDataInsightsClasses returns class names alphabetically,
// ignoring case and breaking ties by the original label for stable worksheets.
func sortedOtherProductsDataInsightsClasses(rowsByClass map[string][]dataInsightsRow) []string {
	classes := make([]string, 0, len(rowsByClass))
	for class := range rowsByClass {
		classes = append(classes, class)
	}
	sort.Slice(classes, func(i, j int) bool {
		left, right := strings.ToUpper(classes[i]), strings.ToUpper(classes[j])
		if left == right {
			return classes[i] < classes[j]
		}
		return left < right
	})
	return classes
}

// sortDataInsightsRows orders occasions alphabetically, without seasonal buckets.
func sortDataInsightsRows(rows []dataInsightsRow) {
	sort.Slice(rows, func(i, j int) bool {
		left, right := strings.ToUpper(rows[i].Occasion), strings.ToUpper(rows[j].Occasion)
		if left == right {
			return rows[i].Occasion < rows[j].Occasion
		}
		return left < right
	})
}

// isExactCounterCards reports whether item belongs to the exact Counter Cards
// class, preferring its unmodified inventory category over the display category.
func isExactCounterCards(item *inventoryEntry) bool {
	category := strings.TrimSpace(item.RawClassDesc)
	if category == "" {
		category = strings.TrimSpace(item.ClassDesc)
	}
	return category == "Counter Cards"
}

// normalizeDataInsightsClassDescription returns item's trimmed original class,
// falling back to its display category and then UNCLASSIFIED when both are empty.
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

// normalizeDataInsightsOccasion returns occ trimmed, or NO OCCASION when empty.
func normalizeDataInsightsOccasion(occ string) string {
	if trimmed := strings.TrimSpace(occ); trimmed != "" {
		return trimmed
	}
	return "NO OCCASION"
}

// totalDataInsightsRows sums a table's inventory, aligned history, and forecasts.
// A total is projectable only when every detail row is projectable, so it cannot
// silently compare a partial forecast with the full table's PY baseline.
func totalDataInsightsRows(rows []dataInsightsRow) dataInsightsRow {
	total := dataInsightsRow{Occasion: "Total", HasProjection: len(rows) > 0}
	for _, row := range rows {
		total.DollarSoldYTD += row.DollarSoldYTD
		total.DollarSoldPY += row.DollarSoldPY
		total.DollarSoldPYToDate += row.DollarSoldPYToDate
		total.ProjectedDollar += row.ProjectedDollar
		total.HasProjection = total.HasProjection && row.HasProjection
		total.LimitedHistory = total.LimitedHistory || row.LimitedHistory
	}
	return total
}
