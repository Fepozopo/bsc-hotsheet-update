package hotsheet

import (
	"fmt"
	"math"
	"strings"
	"time"
)

// occasionDateInfo captures how an occasion should be displayed, ordered, and projected.
type occasionDateInfo struct {
	Display             string
	Month               time.Month
	Day                 int
	SortKey             int
	Complete            bool
	TargetMonthsThrough float64
}

// dataInsightsDateMap maps normalized occasion names to their display text and calendar metadata.
var dataInsightsDateMap = map[string]occasionDateInfo{
	// Valentine's Day cards sell in two separate windows, but we still anchor the occasion to Feb 14
	// so the event keeps its normal calendar identity while projection logic handles the split season.
	"VALENTINE'S DAY":   {Display: "February 14", Month: time.February, Day: 14, SortKey: 214},
	"VALENTINES DAY":    {Display: "February 14", Month: time.February, Day: 14, SortKey: 214},
	"ST PATRICKS DAY":   {Display: "March 17", Month: time.March, Day: 17, SortKey: 317},
	"ST. PATRICK'S DAY": {Display: "March 17", Month: time.March, Day: 17, SortKey: 317},
	"EASTER":            {Display: "April 5", Month: time.April, Day: 4, SortKey: 405},
	"MOTHER'S DAY":      {Display: "May 10", Month: time.May, Day: 10, SortKey: 510},
	"MOTHERS DAY":       {Display: "May 10", Month: time.May, Day: 10, SortKey: 510},
	"GRADUATION":        {Display: "mid-June", Month: time.June, Day: 15, SortKey: 615},
	"FATHER'S DAY":      {Display: "June 21", Month: time.June, Day: 21, SortKey: 621},
	"FATHERS DAY":       {Display: "June 21", Month: time.June, Day: 21, SortKey: 621},
	"INDEPENDENCE DAY":  {Display: "July 4", Month: time.July, Day: 4, SortKey: 704},
	"HALLOWEEN":         {Display: "October 31", Month: time.October, Day: 31, SortKey: 1031},
	"VETERAN'S DAY":     {Display: "November 11", Month: time.November, Day: 11, SortKey: 1111},
	"VETERANS DAY":      {Display: "November 11", Month: time.November, Day: 11, SortKey: 1111},
	"THANKSGIVING":      {Display: "November 26", Month: time.November, Day: 26, SortKey: 1126},
	"HANUKKAH":          {Display: "December 5", Month: time.December, Day: 5, SortKey: 1205},
	"HOLIDAY":           {Display: "December 25", Month: time.December, Day: 25, SortKey: 1225},
	"CHRISTMAS":         {Display: "December 25", Month: time.December, Day: 25, SortKey: 1225},
}

// dataInsightsSeasonTotalFinal derives the section total status text from the rows that
// were actually written so the total matches the same in-progress/complete wording used by
// the detail rows above it.
func dataInsightsSeasonTotalFinal(totalProjected, totalPY float64, rows []dataInsightsRow) string {
	sectionStarted := true
	sectionComplete := true
	for _, row := range rows {
		if strings.HasPrefix(row.Final, "NOT STARTED") {
			sectionStarted = false
			sectionComplete = false
		}
		if strings.HasPrefix(row.Final, "IN PROGRESS:") {
			sectionComplete = false
		}
	}
	if !sectionStarted {
		return formatSeasonStatusYoY(totalProjected, totalPY, false, false)
	}
	return formatSeasonStatusYoY(totalProjected, totalPY, true, sectionComplete)
}

// projectDataInsightsRow centralizes the seasonal projection rules so both card and
// non-card Data Insights rows use the same date metadata and year-over-year logic.
func projectDataInsightsRow(section, occasion string, dollarSoldYTD, dollarSoldPY, currentMonthsThrough float64, now time.Time) (occasionDateInfo, float64, string) {
	dateInfo := dataInsightsDateInfo(section, occasion, now)

	if section == "Spring" || section == "Winter" {
		// Seasonal items are projected only up to their typical selling window.
		if isValentinesOccasion(occasion) {
			// Valentine's Day is the one split-season exception: it sells in an early-year
			// window and again near the end of the year, so we project against both windows.
			currentSellingDays, totalSellingDays, complete := valentinesProjectionWindow(now)
			projected := dollarSoldYTD * (totalSellingDays / currentSellingDays)
			return dateInfo, projected, formatSeasonStatusYoY(projected, dollarSoldPY, true, complete)
		}
		if section == "Winter" {
			// Winter items are treated as selling from July 1 forward, which matches the
			// existing card logic and avoids projecting winter holidays from the off-season.
			seasonStart := time.Date(now.Year(), time.July, 1, 0, 0, 0, 0, now.Location())
			if now.Before(seasonStart) {
				projected := dollarSoldYTD
				return dateInfo, projected, formatSeasonStatusYoY(projected, dollarSoldPY, false, dateInfo.Complete)
			}
			if dateInfo.Complete {
				projected := dollarSoldYTD
				return dateInfo, projected, formatSeasonStatusYoY(projected, dollarSoldPY, true, dateInfo.Complete)
			}
			currentSellingMonths := monthsThroughSinceDate(now.Year(), time.July, 1, now.Month(), now.Day(), now.Location())
			projected := dollarSoldYTD * (dateInfo.TargetMonthsThrough / currentSellingMonths)
			return dateInfo, projected, formatSeasonStatusYoY(projected, dollarSoldPY, true, dateInfo.Complete)
		}
		if dateInfo.Complete {
			// Once the holiday has passed, stop extrapolating and keep the row at actual YTD.
			projected := dollarSoldYTD
			return dateInfo, projected, formatSeasonStatusYoY(projected, dollarSoldPY, true, dateInfo.Complete)
		}
		projected := dollarSoldYTD * (dateInfo.TargetMonthsThrough / currentMonthsThrough)
		return dateInfo, projected, formatSeasonStatusYoY(projected, dollarSoldPY, true, dateInfo.Complete)
	}

	// Everyday items are projected across the full year because they do not have a holiday
	// cutoff date to anchor them to.
	projected := dollarSoldYTD * (12.0 / currentMonthsThrough)
	return dateInfo, projected, formatYoYFromProjectedSales(projected, dollarSoldPY)
}

// dataInsightsDateInfo returns the display and projection metadata for a data-insight occasion.
func dataInsightsDateInfo(section, occasion string, now time.Time) occasionDateInfo {
	if section == "Everyday" {
		return occasionDateInfo{Display: "N/A", SortKey: 999999, TargetMonthsThrough: 12.0}
	}

	// Unknown seasonal occasions fall back to a neutral display and sort after known holidays.
	info, ok := dataInsightsDateMap[strings.ToUpper(strings.TrimSpace(occasion))]
	if !ok {
		return occasionDateInfo{Display: "N/A", SortKey: 999999, TargetMonthsThrough: 12.0}
	}

	if isValentinesOccasion(occasion) {
		// Valentine's Day is the one seasonal occasion that is sold in two merchandising waves:
		// an early-year run through Feb 14, then a late-year run from Nov 16 through Dec 31.
		// We show that split directly in the sheet and keep it marked incomplete until year end,
		// because December inventory is still part of the same selling season.
		info.Display = "Jan 1 - Feb 14, Nov 16 - Dec 31"
		info.Complete = !now.Before(time.Date(now.Year()+1, time.January, 1, 0, 0, 0, 0, now.Location()))
		info.TargetMonthsThrough = monthsThroughForDate(now.Year(), info.Month, info.Day, now.Location())
		return info
	}

	if section == "Winter" {
		// Winter holidays don't really start selling until July 1, so treat that as the
		// beginning of the season when calculating the projection window.
		info.TargetMonthsThrough = monthsThroughSinceDate(now.Year(), time.July, 1, info.Month, info.Day, now.Location())
	} else {
		info.TargetMonthsThrough = monthsThroughForDate(now.Year(), info.Month, info.Day, now.Location())
	}

	// Treat the occasion as complete for the whole event day, not just after midnight.
	eventDate := time.Date(now.Year(), info.Month, info.Day, 23, 59, 59, 0, now.Location())
	info.Complete = !now.Before(eventDate)
	return info
}

// isValentinesOccasion centralizes the occasion matching so both spelling variants
// take the same split-window projection path.
func isValentinesOccasion(occasion string) bool {
	switch strings.ToUpper(strings.TrimSpace(occasion)) {
	case "VALENTINE'S DAY", "VALENTINES DAY":
		return true
	default:
		return false
	}
}

// valentinesProjectionWindow returns the active and total selling-day counts for the
// split Valentine's season so projections can scale against the portion of the year
// that is actually on sale instead of assuming the occasion sells all year.
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

// monthsThroughSinceDate returns the month progress between two calendar dates.
func monthsThroughSinceDate(year int, startMonth time.Month, startDay int, endMonth time.Month, endDay int, loc *time.Location) float64 {
	monthsThrough := monthsThroughForDate(year, endMonth, endDay, loc) - monthsThroughForDate(year, startMonth, startDay, loc)
	if monthsThrough <= 0 {
		monthsThrough = 1
	}
	return monthsThrough
}

// formatSeasonStatusYoY formats the season status text with a YoY comparison.
func formatSeasonStatusYoY(projectedSales float64, pySales float64, started bool, complete bool) string {
	if !started {
		return fmt.Sprintf("NOT STARTED: %s YoY", formatYoYFromProjectedSales(projectedSales, pySales))
	}
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
