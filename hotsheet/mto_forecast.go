package hotsheet

import (
	"fmt"
	"sort"
	"strings"
	"time"
)

const (
	mtoForecastMonths = 12
	mtoHorizonMonths  = 24
	mtoHistoryYears   = 3
)

// mtoForecastState distinguishes an estimated stockout from a longer runway or
// insufficient complete sales history; it controls both display and row order.
type mtoForecastState uint8

const (
	mtoStockout mtoForecastState = iota
	mtoBeyondHorizon
	mtoInsufficientHistory
)

// mtoForecastRow contains an active inventory item's BSC sales forecast and
// stockout estimate, or an explicit reason why no date can be estimated.
type mtoForecastRow struct {
	item      *inventoryEntry
	available int
	demand12  float64
	monthly   [12]float64
	hasDemand bool
	mto       float64
	stockout  string
	coverage  string
	state     mtoForecastState
}

// observedMonth records one completed month in the history export. The year is
// used to favor newer observations of the same calendar month.
type observedMonth struct {
	year  int
	units float64
}

// mtoHistoryProfile holds forecast units per full calendar month and the number
// of completed observations since the item's first positive BSC sale.
type mtoHistoryProfile struct {
	monthly [12]float64
	months  int
	years   int
	usable  bool
}

// buildMTOHistoryProfile builds a seasonal profile from BSC Quantity Sold records
// completed before asOf. Months before the first positive sale and the current
// partial month are excluded; completed zero months after that sale are included.
// It requires at least one observation for every calendar month before forecasting.
func buildMTOHistoryProfile(records []salesRecord, asOf time.Time) mtoHistoryProfile {
	var profile mtoHistoryProfile
	firstSale := 0
	foundSale := false
	for _, record := range records {
		if record.Metric != "Quantity Sold" {
			continue
		}
		for index, units := range record.Periods {
			month := time.Month(index + 1)
			if units > 0 && !time.Date(record.Year, month+1, 1, 0, 0, 0, 0, time.UTC).After(asOf) {
				key := record.Year*12 + index
				if !foundSale || key < firstSale {
					firstSale = key
					foundSale = true
				}
			}
		}
	}
	if !foundSale {
		return profile
	}

	var observed [12][]observedMonth
	seenYears := make(map[int]struct{})
	for _, record := range records {
		if record.Metric != "Quantity Sold" {
			continue
		}
		for index, units := range record.Periods {
			month := time.Month(index + 1)
			if record.Year*12+index < firstSale || time.Date(record.Year, month+1, 1, 0, 0, 0, 0, time.UTC).After(asOf) {
				continue
			}
			// Sold units represent demand; corrections cannot create negative
			// future depletion. Zero sales after the first sale remain observed.
			observed[index] = append(observed[index], observedMonth{year: record.Year, units: max(units, 0)})
			profile.months++
			seenYears[record.Year] = struct{}{}
		}
	}
	profile.years = len(seenYears)
	profile.usable = true
	for index := range observed {
		months := observed[index]
		if len(months) == 0 {
			profile.usable = false
			continue
		}
		sort.Slice(months, func(i, j int) bool { return months[i].year > months[j].year })
		if len(months) > mtoHistoryYears {
			months = months[:mtoHistoryYears]
		}
		var weighted, totalWeight float64
		for i, month := range months {
			weight := float64(len(months) - i)
			weighted += month.units * weight
			totalWeight += weight
		}
		profile.monthly[index] = weighted / totalWeight
	}
	return profile
}

// forecastMTOWindow projects demand between start and end by prorating each
// calendar month's seasonal units over the days in that month. It also reports
// the first month in which available units are exhausted and elapsed months;
// within-month demand is assumed uniform because daily sales are unavailable.
func forecastMTOWindow(monthly [12]float64, start, end time.Time, available float64) (demand, monthsTillOut float64, stockout string) {
	remaining := available
	elapsed := 0.0
	for monthStart := time.Date(start.Year(), start.Month(), 1, 0, 0, 0, 0, time.UTC); monthStart.Before(end); monthStart = monthStart.AddDate(0, 1, 0) {
		monthEnd := monthStart.AddDate(0, 1, 0)
		segmentStart, segmentEnd := monthStart, monthEnd
		if segmentStart.Before(start) {
			segmentStart = start
		}
		if segmentEnd.After(end) {
			segmentEnd = end
		}
		fraction := segmentEnd.Sub(segmentStart).Hours() / monthEnd.Sub(monthStart).Hours()
		units := monthly[int(monthStart.Month())-1]
		segmentDemand := units * fraction
		demand += segmentDemand
		if stockout == "" && units > 0 && remaining <= segmentDemand {
			stockout = monthStart.Format("Jan 2006")
			monthsTillOut = elapsed + remaining/units
		}
		remaining -= segmentDemand
		elapsed += fraction
	}
	return demand, monthsTillOut, stockout
}

// monthsAfter returns the same day of month after months calendar months, clamping
// to the target month's last day (for example, February 29 to February 28).
func monthsAfter(start time.Time, months int) time.Time {
	first := time.Date(start.Year(), start.Month()+time.Month(months), 1, 0, 0, 0, 0, time.UTC)
	day := min(start.Day(), first.AddDate(0, 1, -1).Day())
	return first.AddDate(0, 0, day-1)
}

// buildMTORows forecasts active inventory entries as of the report date, including
// undated POs in available stock and only BSC units sold in demand. It returns
// rows ordered by earliest stockout, then >24-month and insufficient-history items.
func buildMTORows(entries []*inventoryEntry, asOf time.Time) []mtoForecastRow {
	rows := make([]mtoForecastRow, 0, len(entries))
	// Inventory quantities are a report-date snapshot; project from the next day.
	start := asOf.AddDate(0, 0, 1)
	end12, end24 := monthsAfter(start, mtoForecastMonths), monthsAfter(start, mtoHorizonMonths)
	for _, item := range entries {
		switch strings.ToLower(strings.ReplaceAll(strings.TrimSpace(item.Status), " ", "")) {
		case "rundown", "discontinued":
			continue
		}
		row := mtoForecastRow{item: item, available: item.OnHand + item.OnPO - item.OnSO - item.OnBO}
		profile := buildMTOHistoryProfile(item.SalesRecords, asOf)
		row.coverage = fmt.Sprintf("%d months / %d years", profile.months, profile.years)
		if profile.usable {
			row.monthly = profile.monthly
			row.demand12, _, _ = forecastMTOWindow(profile.monthly, start, end12, float64(row.available))
			row.hasDemand = true
		}
		switch {
		case row.available <= 0:
			row.state = mtoStockout
			row.stockout = asOf.Format("Jan 2006")
		case !profile.usable:
			row.state = mtoInsufficientHistory
			row.stockout = "Insufficient history"
		default:
			_, row.mto, row.stockout = forecastMTOWindow(profile.monthly, start, end24, float64(row.available))
			if row.stockout == "" {
				row.state = mtoBeyondHorizon
				row.stockout = fmt.Sprintf("Not within %d months", mtoHorizonMonths)
			}
		}
		rows = append(rows, row)
	}
	sort.Slice(rows, func(i, j int) bool {
		left, right := rows[i], rows[j]
		if left.state != right.state {
			return left.state < right.state
		}
		if left.state == mtoStockout && left.mto != right.mto {
			return left.mto < right.mto
		}
		return left.item.SKU < right.item.SKU
	})
	return rows
}
