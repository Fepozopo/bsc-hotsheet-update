package hotsheet

import (
	"fmt"
	"sort"
	"strings"
	"time"
)

const (
	mtoForecastMonths   = 12
	mtoHorizonMonths    = 24
	mtoHistoryYears     = 3
	mtoMinHistoryMonths = 2

	mtoKeyAccountHistoryStartYear  = 2026
	mtoKeyAccountHistoryStartMonth = time.February
)

// mtoForecastState distinguishes an estimated stockout from a longer runway or
// insufficient completed shipment history; it controls display and row order.
type mtoForecastState uint8

const (
	mtoStockout mtoForecastState = iota
	mtoBeyondHorizon
	mtoInsufficientHistory
)

// mtoForecastRow contains an active inventory item's BSC shipment forecast and
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

// observedMonth records one completed month of shipped units in the history
// exports. The year favors newer observations of the same calendar month.
type observedMonth struct {
	year  int
	units float64
}

// mtoHistoryProfile holds forecast units per full calendar month and the number
// of completed observations since the item's first positive BSC shipment.
type mtoHistoryProfile struct {
	monthly [12]float64
	months  int
	years   int
	usable  bool
}

// shippedRecords combines each calendar month's sales and nonnegative issued
// quantities into one BSC Quantity Shipped record per year. Other history metrics
// are ignored so forecast coverage counts each month at most once.
func shippedRecords(records []salesRecord) []salesRecord {
	byYear := make(map[int]*salesRecord)
	for _, record := range records {
		if record.Metric != "Quantity Sold" && record.Metric != "Quantity Issued" {
			continue
		}
		shipped := byYear[record.Year]
		if shipped == nil {
			shipped = &salesRecord{Year: record.Year, Metric: "Quantity Shipped"}
			byYear[record.Year] = shipped
		}
		for month, units := range record.Periods {
			if record.Metric == "Quantity Issued" {
				units = max(units, 0)
			}
			shipped.Periods[month] += units
		}
	}
	years := make([]int, 0, len(byYear))
	for year := range byYear {
		years = append(years, year)
	}
	sort.Ints(years)
	shipped := make([]salesRecord, 0, len(years))
	for _, year := range years {
		shipped = append(shipped, *byYear[year])
	}
	return shipped
}

// buildMTOHistoryProfile forecasts monthly BSC shipped units completed by asOf.
// Months before the first positive shipment and the current partial month are
// excluded; completed zero months afterward count once. With at least two
// observations, missing calendar months use the observed mean while measured
// months retain recent-year weighted estimates.
func buildMTOHistoryProfile(records []salesRecord, asOf time.Time) mtoHistoryProfile {
	var profile mtoHistoryProfile
	shipments := shippedRecords(records)
	firstSale := 0
	foundSale := false
	for _, record := range shipments {
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
	var totalUnits float64
	seenYears := make(map[int]struct{})
	for _, record := range shipments {
		for index, units := range record.Periods {
			month := time.Month(index + 1)
			if record.Year*12+index < firstSale || time.Date(record.Year, month+1, 1, 0, 0, 0, 0, time.UTC).After(asOf) {
				continue
			}
			// Shipped units represent demand; corrections cannot create negative
			// future depletion. Zero shipments after the first remain observed.
			units = max(units, 0)
			observed[index] = append(observed[index], observedMonth{year: record.Year, units: units})
			totalUnits += units
			profile.months++
			seenYears[record.Year] = struct{}{}
		}
	}
	profile.years = len(seenYears)
	profile.usable = profile.months >= mtoMinHistoryMonths
	if !profile.usable {
		return profile
	}
	// Without a same-calendar-month observation, a cross-month mean supplies a
	// provisional estimate while leaving measured seasonal months untouched.
	fallback := totalUnits / float64(profile.months)
	for index := range observed {
		months := observed[index]
		if len(months) == 0 {
			profile.monthly[index] = fallback
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

// mtoHistoryRecords returns the item's history eligible for MTO forecasting.
// Listed product-line 2021 SKUs exclude history before February 2026, when the
// key account moved to custom SKUs. Excluded periods in the cutoff year are
// zeroed in forecast-only copies. The source history is never changed so other
// sheets retain their existing sales and issue data.
func mtoHistoryRecords(item *inventoryEntry) []salesRecord {
	if strings.TrimSpace(item.ProductLine) != "2021" {
		return item.SalesRecords
	}
	switch strings.TrimSpace(item.SKU) {
	case "BD1001FJ", "BD1006F", "BD1028", "BD1044", "BD1046F", "BD1048FJ",
		"BP1002", "CS1003", "CS1005", "FC1005", "GR1016FJ", "HY1048FB",
		"HY1071", "HY1071B", "LV1010", "MD1011", "MI1011", "MI1014",
		"PJ1011", "PJ1031", "PJ1032", "RL1003", "SP1001F", "TY1007",
		"TY1019", "TY1019B", "TY1020", "TY1023", "TY1023B", "TY1025", "TY1025B",
		"TY1028F", "VD1029J":
		// Copy eligible records rather than filtering in place: the shared
		// slice also feeds Monthly History and range-based Best Sellers.
		records := make([]salesRecord, 0, len(item.SalesRecords))
		for _, record := range item.SalesRecords {
			if record.Year < mtoKeyAccountHistoryStartYear {
				continue
			}
			if record.Year == mtoKeyAccountHistoryStartYear {
				// Excluded shipments must not establish the first positive month.
				// The profile then omits these pre-sale zeros from coverage too.
				clear(record.Periods[:mtoKeyAccountHistoryStartMonth-1])
			}
			records = append(records, record)
		}
		return records
	default:
		return item.SalesRecords
	}
}

// buildMTORows forecasts active inventory entries as of the report date, including
// undated POs in available stock and subtracting committed quantity (sales orders
// plus back orders). BSC shipped units count as demand; listed product-line 2021
// SKUs use only February-2026-and-later history. It returns rows ordered by earliest
// stockout, then >24-month and insufficient-history items.
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
		committed := item.OnSO + item.OnBO
		row := mtoForecastRow{item: item, available: item.OnHand + item.OnPO - committed}
		profile := buildMTOHistoryProfile(mtoHistoryRecords(item), asOf)
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
