package hotsheet

import (
	"fmt"
	"math"
	"sort"
	"strings"
	"time"
)

const (
	dataInsightsHistoryYears = 3
	dataInsightsMinMonths    = 2
	dataInsightsNoPYBaseline = "N/A — no PY baseline"
	dataInsightsInsufficient = "Insufficient history"
)

// dataInsightsHolidayLabels gives fixed holidays stable dates and movable
// holidays names instead of dates that become incorrect in another report year.
var dataInsightsHolidayLabels = map[string]string{
	"VALENTINE'S DAY": "Feb 14", "VALENTINES DAY": "Feb 14",
	"ST PATRICKS DAY": "Mar 17", "ST. PATRICK'S DAY": "Mar 17",
	"EASTER": "Easter", "PASSOVER": "Passover",
	"MOTHER'S DAY": "Mother's Day", "MOTHERS DAY": "Mother's Day",
	"FATHER'S DAY": "Father's Day", "FATHERS DAY": "Father's Day",
	"INDEPENDENCE DAY": "Jul 4", "HALLOWEEN": "Oct 31",
	"VETERAN'S DAY": "Nov 11", "VETERANS DAY": "Nov 11",
	"THANKSGIVING": "Thanksgiving", "HANUKKAH": "Hanukkah",
	"HOLIDAY": "Dec 25", "CHRISTMAS": "Dec 25",
}

// dataInsightsHolidayLabel returns a stable holiday/date label for occasion.
// Undated occasions, including Graduation and Everyday, return N/A.
func dataInsightsHolidayLabel(occasion string) string {
	if label, ok := dataInsightsHolidayLabels[strings.ToUpper(strings.TrimSpace(occasion))]; ok {
		return label
	}
	return "N/A"
}

// projectDataInsightsRow fills row's full-year dollar estimate from aggregate
// history as of runDate. Only Dollars Sold records supplied by the row's SKUs
// are aggregated; missing SKU/year records never establish zero observations.
// PY-to-date and current-year actuals use January through the last completed
// month. Forecast observations start at the combined row's first positive sale,
// omit the current/future months, and retain at most three recent observations
// per calendar month. Signed dollar values are retained, unlike unit demand.
func projectDataInsightsRow(row *dataInsightsRow, history map[int][12]float64, runDate time.Time) {
	completed := int(runDate.Month()) - 1
	year := runDate.Year()
	for month := 0; month < completed; month++ {
		row.DollarSoldPYToDate += history[year-1][month]
	}

	years := make([]int, 0, len(history))
	for historyYear := range history {
		if historyYear <= year {
			years = append(years, historyYear)
		}
	}
	sort.Ints(years)
	// Leading zeros do not establish a launch date. Once positive aggregate
	// sales exist, subsequent completed zeros in supplied records are observed.
	firstSale := 0
	foundSale := false
	for _, historyYear := range years {
		for month, dollars := range history[historyYear] {
			if historyYear == year && month >= completed {
				break
			}
			if dollars > 0 {
				firstSale = historyYear*12 + month
				foundSale = true
				break
			}
		}
		if foundSale {
			break
		}
	}
	if !foundSale {
		return
	}

	var monthly [12]float64
	var observed [12]bool
	var observations, calendarMonths int
	var total float64
	for month := range monthly {
		var recent [dataInsightsHistoryYears]float64
		count := 0
		for index := len(years) - 1; index >= 0 && count < len(recent); index-- {
			historyYear := years[index]
			if historyYear*12+month < firstSale || (historyYear == year && month >= completed) {
				continue
			}
			recent[count] = history[historyYear][month]
			count++
		}
		if count == 0 {
			continue
		}
		observed[month] = true
		calendarMonths++
		observations += count
		var weighted, weights float64
		for index, dollars := range recent[:count] {
			// Match MTO's recency ranks: 3:2:1 with three comparable years,
			// 2:1 with two, and no dilution by absent SKU/year records.
			weight := float64(count - index)
			weighted += dollars * weight
			weights += weight
			total += dollars
		}
		monthly[month] = weighted / weights
	}
	if observations < dataInsightsMinMonths {
		return
	}

	// A missing calendar month receives the unweighted observed-month mean,
	// not zero. Observed zero months retain their own weighted estimates.
	fallback := total / float64(observations)
	for month := range monthly {
		if !observed[month] {
			monthly[month] = fallback
		}
		if month < completed {
			row.ProjectedDollar += history[year][month]
		} else {
			row.ProjectedDollar += monthly[month]
		}
	}
	row.HasProjection = true
	row.LimitedHistory = calendarMonths < len(monthly)
}

// dataInsightsYoYDisplay returns the full-year comparison for row, prioritizing
// a missing PY baseline over history sufficiency. Limited calendar coverage is
// disclosed alongside the percentage, never as a seasonal status.
func dataInsightsYoYDisplay(row dataInsightsRow) string {
	if row.DollarSoldPY == 0 {
		return dataInsightsNoPYBaseline
	}
	if !row.HasProjection {
		return dataInsightsInsufficient
	}
	pct := math.Round((row.ProjectedDollar/row.DollarSoldPY - 1) * 100)
	if pct == 0 {
		// Small declines can round to IEEE negative zero; an unchanged
		// whole-percentage outlook should not display a misleading -0%.
		pct = 0
	}
	display := fmt.Sprintf("%+.0f%%", pct)
	if row.LimitedHistory {
		display += " (limited-history estimate)"
	}
	return display
}
