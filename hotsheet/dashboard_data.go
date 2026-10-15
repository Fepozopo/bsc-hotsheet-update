package hotsheet

import (
	"fmt"
	"sort"
	"time"
)

// dashboardTopItems is the number of ranked slices reserved for each donut.
const dashboardTopItems = 15

// dashboardItem is a labeled, positive donut value; its zero value pads unused ranks.
type dashboardItem struct {
	label string
	value float64
}

// dashboardClassData contains chart values for a unique class display label.
// Sales and monthly index 0 is the report year, index 1 is the prior year;
// monthly indexes 0–11 represent January–December. Donuts use inventory YTD.
type dashboardClassData struct {
	class   string
	shipped [dashboardTopItems]dashboardItem
	dollars [dashboardTopItems]dashboardItem
	sales   [2]float64
	monthly [2][12]float64
}

// buildDashboardData returns All followed by alphabetically sorted class display
// rows, including every status and ignoring nil entries. Classes prefer RawClassDesc,
// then ClassDesc, then Unclassified. A real All class uses All (class), or a numbered
// variant if that label is already a class name, so it never merges with the aggregate.
// With history, sales include only completed months in runDate's year and the prior
// year; otherwise they use inventory dollar totals. Monthly values always sum full
// report-year and prior-year Dollars Sold records. Inputs are never modified, and
// runDate is the only date source.
func buildDashboardData(entries []*inventoryEntry, hasHistory bool, runDate time.Time) []dashboardClassData {
	byClass := make(map[string][]*inventoryEntry)
	all := make([]*inventoryEntry, 0, len(entries))
	for _, entry := range entries {
		if entry == nil {
			continue
		}
		class := entry.RawClassDesc
		if class == "" {
			class = entry.ClassDesc
		}
		if class == "" {
			class = "Unclassified"
		}
		byClass[class] = append(byClass[class], entry)
		all = append(all, entry)
	}

	rows := make([]dashboardClassData, 1, len(byClass)+1)
	rows[0] = summarizeDashboardClass(all, hasHistory, runDate)
	rows[0].class = "All"
	for class, items := range byClass {
		row := summarizeDashboardClass(items, hasHistory, runDate)
		row.class = class
		if class == "All" {
			// Reserve original class names before choosing a synthetic dropdown label.
			row.class = "All (class)"
			for suffix := 2; ; suffix++ {
				if _, exists := byClass[row.class]; !exists {
					break
				}
				row.class = fmt.Sprintf("All (class %d)", suffix)
			}
		}
		rows = append(rows, row)
	}
	sort.Slice(rows[1:], func(i, j int) bool { return rows[i+1].class < rows[j+1].class })
	return rows
}

// summarizeDashboardClass returns rankings and dollar series for non-nil entries.
// hasHistory selects completed-month sales instead of inventory totals; runDate
// anchors both history years and the exclusive current-month cutoff.
func summarizeDashboardClass(entries []*inventoryEntry, hasHistory bool, runDate time.Time) dashboardClassData {
	data := dashboardClassData{
		shipped: rankDashboardItems(entries, true),
		dollars: rankDashboardItems(entries, false),
	}
	completedMonths := int(runDate.Month()) - 1
	for _, entry := range entries {
		if !hasHistory {
			data.sales[0] += entry.DollarSoldYTD
			data.sales[1] += entry.DollarSoldPY
		}
		for _, record := range entry.SalesRecords {
			yearIndex := runDate.Year() - record.Year
			if record.Metric != "Dollars Sold" || yearIndex < 0 || yearIndex >= len(data.monthly) {
				continue
			}
			for month, value := range record.Periods {
				data.monthly[yearIndex][month] += value
				if hasHistory && month < completedMonths {
					data.sales[yearIndex] += value
				}
			}
		}
	}
	return data
}

// rankDashboardItems returns up to fifteen positive inventory values, descending
// by value then ascending by SKU, with unused ranks zero-padded. shipped selects
// YTDSold plus nonnegative YTDIssued; otherwise DollarSoldYTD is ranked. Labels
// append the description only when present, and sorting never changes entries.
func rankDashboardItems(entries []*inventoryEntry, shipped bool) [dashboardTopItems]dashboardItem {
	// candidate pairs an entry with its ranking value, keeping the SKU separate
	// from the display label so descriptions cannot break ties.
	type candidate struct {
		entry *inventoryEntry
		value float64
	}
	candidates := make([]candidate, 0, len(entries))
	for _, entry := range entries {
		value := entry.DollarSoldYTD
		if shipped {
			value = float64(entry.YTDSold) + float64(max(entry.YTDIssued, 0))
		}
		if value > 0 {
			candidates = append(candidates, candidate{entry: entry, value: value})
		}
	}
	sort.SliceStable(candidates, func(i, j int) bool {
		if candidates[i].value == candidates[j].value {
			return candidates[i].entry.SKU < candidates[j].entry.SKU
		}
		return candidates[i].value > candidates[j].value
	})
	var ranked [dashboardTopItems]dashboardItem
	for rank, candidate := range candidates[:min(len(candidates), len(ranked))] {
		label := candidate.entry.SKU
		if candidate.entry.Description != "" {
			label += " - " + candidate.entry.Description
		}
		ranked[rank] = dashboardItem{label: label, value: candidate.value}
	}
	return ranked
}
