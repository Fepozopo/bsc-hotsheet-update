package hotsheet

import (
	"sort"
	"strings"
	"time"
)

// dataInsightsRow holds an occasion or whole table's inventory totals and history-based
// outlook. DollarSoldYTD and DollarSoldPY always retain inventory-report values;
// DollarSoldPYToDate and ProjectedDollar use optional monthly dollar history.
// HasProjection distinguishes a real zero estimate from unavailable history.
// Class names each Other Products detail row's table; card rows and totals leave it empty.
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

// dataInsightsTable contains occasion rows and an independently forecast total.
// The total uses combined monthly history, not the sum of occasion forecasts.
type dataInsightsTable struct {
	Rows  []dataInsightsRow
	Total dataInsightsRow
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

// buildDataInsightsTables returns Counter Cards grouped by occasion and Other
// Products grouped by class then occasion, each with an independently forecast
// total. Labels are trimmed and grouped case-insensitively; rows are sorted here
// and classes by the writer. Nil entries are ignored. hasHistory enables
// projections anchored exclusively to runDate; without history, inventory
// totals remain but projections do not.
func buildDataInsightsTables(entries []*inventoryEntry, hasHistory bool, runDate time.Time) (dataInsightsTable, map[string]dataInsightsTable) {
	groups := make(map[dataInsightsGroupKey]*dataInsightsGroup)
	groupsByClass := make(map[string][]*dataInsightsGroup)
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
			groupsByClass[class] = append(groupsByClass[class], group)
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

	cards := dataInsightsTable{Total: dataInsightsRow{Occasion: "Total"}}
	other := make(map[string]dataInsightsTable)
	for class, classGroups := range groupsByClass {
		table := dataInsightsTable{Total: totalDataInsightsRows(classGroups, hasHistory, runDate)}
		for _, group := range classGroups {
			if hasHistory && !runDate.IsZero() {
				projectDataInsightsRow(&group.row, group.history, runDate)
			}
			table.Rows = append(table.Rows, group.row)
		}
		sortDataInsightsRows(table.Rows)
		if class == "" {
			cards = table
		} else {
			other[class] = table
		}
	}
	return cards, other
}

// sortedOtherProductsDataInsightsClasses returns table class names alphabetically,
// ignoring case and breaking ties by the original label for stable worksheets.
func sortedOtherProductsDataInsightsClasses(tablesByClass map[string]dataInsightsTable) []string {
	classes := make([]string, 0, len(tablesByClass))
	for class := range tablesByClass {
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

// totalDataInsightsRows sums groups' inventory values and observed monthly
// dollars, then forecasts that combined history as of runDate when hasHistory
// is true. The total applies its own first-positive-sale and two-month rules;
// an unprojectable occasion cannot veto an otherwise usable table history.
// Missing SKU/year records remain absent, and source groups are not modified.
func totalDataInsightsRows(groups []*dataInsightsGroup, hasHistory bool, runDate time.Time) dataInsightsRow {
	total := dataInsightsRow{Occasion: "Total"}
	history := make(map[int][12]float64)
	for _, group := range groups {
		total.DollarSoldYTD += group.row.DollarSoldYTD
		total.DollarSoldPY += group.row.DollarSoldPY
		for year, dollars := range group.history {
			combined := history[year]
			for month, value := range dollars {
				combined[month] += value
			}
			history[year] = combined
		}
	}
	if hasHistory && !runDate.IsZero() {
		projectDataInsightsRow(&total, history, runDate)
	}
	return total
}
