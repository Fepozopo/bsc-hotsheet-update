package hotsheet

import (
	"math"
	"strings"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// TestValentinesProjectionWindow verifies the split-window selling-day math used for
// Valentine's Day projections.
func TestValentinesProjectionWindow(t *testing.T) {
	t.Parallel()

	now := time.Date(2026, time.March, 1, 12, 0, 0, 0, time.UTC)
	current, total, complete := valentinesProjectionWindow(now)

	if current != 45 {
		t.Fatalf("expected 45 current selling days before the November window, got %v", current)
	}
	if total != 91 {
		t.Fatalf("expected 91 total selling days across both Valentine windows, got %v", total)
	}
	if complete {
		t.Fatal("expected Valentine's Day to remain incomplete before year end")
	}

	current, total, complete = valentinesProjectionWindow(time.Date(2026, time.December, 31, 12, 0, 0, 0, time.UTC))
	if current != total {
		t.Fatalf("expected current selling days to equal the full window on Dec 31, got %v vs %v", current, total)
	}
	if complete {
		t.Fatal("expected Valentine's Day to remain incomplete on Dec 31")
	}
}

// TestBuildDataInsightsRowsValentinesProjection confirms the card section keeps the split
// Valentine's Day date and projection behavior.
func TestBuildDataInsightsRowsValentinesProjection(t *testing.T) {
	t.Parallel()

	entries := []*entry{{
		RawClassDesc:  "Counter Cards",
		Occasion:      "Valentine's Day",
		DollarSoldYTD: 100,
		DollarSoldPY:  80,
	}}

	now := time.Date(2026, time.March, 1, 12, 0, 0, 0, time.UTC)
	rowsBySection := buildDataInsightsRows(entries, currentMonthsThrough(now), now)
	rows := rowsBySection["Spring"]
	if len(rows) != 1 {
		t.Fatalf("expected one Spring row, got %d", len(rows))
	}

	row := rows[0]
	if row.Date != "Jan 1 - Feb 14, Nov 16 - Dec 31" {
		t.Fatalf("expected split-window Valentine's Day display date, got %q", row.Date)
	}

	expectedProjected := 100.0 * 91.0 / 45.0
	if diff := math.Abs(row.ProjectedDollar - expectedProjected); diff > 1e-9 {
		t.Fatalf("expected projected sales %.9f, got %.9f", expectedProjected, row.ProjectedDollar)
	}
	if !strings.HasPrefix(row.Final, "IN PROGRESS:") {
		t.Fatalf("expected in-progress status before year end, got %q", row.Final)
	}

	rowsBySection = buildDataInsightsRows(entries, currentMonthsThrough(time.Date(2026, time.December, 31, 12, 0, 0, 0, time.UTC)), time.Date(2026, time.December, 31, 12, 0, 0, 0, time.UTC))
	row = rowsBySection["Spring"][0]
	if diff := math.Abs(row.ProjectedDollar - 100.0); diff > 1e-9 {
		t.Fatalf("expected projected sales to match YTD at the end of the season, got %.9f", row.ProjectedDollar)
	}
	if !strings.HasPrefix(row.Final, "IN PROGRESS:") {
		t.Fatalf("expected status to remain in progress on Dec 31, got %q", row.Final)
	}
}

// TestGraduationProjectionWindow confirms a normal spring holiday uses its holiday date
// and flips to complete after the event day.
func TestGraduationProjectionWindow(t *testing.T) {
	t.Parallel()

	entries := []*entry{{
		RawClassDesc:  "Counter Cards",
		Occasion:      "Graduation",
		DollarSoldYTD: 100,
		DollarSoldPY:  80,
	}}

	beforeCutoff := time.Date(2026, time.June, 14, 12, 0, 0, 0, time.UTC)
	rowsBySection := buildDataInsightsRows(entries, currentMonthsThrough(beforeCutoff), beforeCutoff)
	rows := rowsBySection["Spring"]
	if len(rows) != 1 {
		t.Fatalf("expected one Spring row, got %d", len(rows))
	}

	row := rows[0]
	if row.Date != "mid-June" {
		t.Fatalf("expected Graduation to display as mid-June, got %q", row.Date)
	}
	if !strings.HasPrefix(row.Final, "IN PROGRESS:") {
		t.Fatalf("expected Graduation to be in progress before June 15, got %q", row.Final)
	}

	afterCutoff := time.Date(2026, time.June, 16, 12, 0, 0, 0, time.UTC)
	rowsBySection = buildDataInsightsRows(entries, currentMonthsThrough(afterCutoff), afterCutoff)
	row = rowsBySection["Spring"][0]
	if row.Date != "mid-June" {
		t.Fatalf("expected Graduation to still display as mid-June, got %q", row.Date)
	}
	if !strings.HasPrefix(row.Final, "COMPLETE:") {
		t.Fatalf("expected Graduation to be complete after June 15, got %q", row.Final)
	}
}

// TestBuildDataInsightsRowsWinterProjectionStartsJuly1 confirms winter holidays stay at
// YTD before July 1 and switch to in-season projection afterward.
func TestBuildDataInsightsRowsWinterProjectionStartsJuly1(t *testing.T) {
	t.Parallel()

	entries := []*entry{{
		RawClassDesc:  "Counter Cards",
		Occasion:      "Christmas",
		DollarSoldYTD: 100,
		DollarSoldPY:  80,
	}}

	beforeSeason := time.Date(2026, time.June, 1, 12, 0, 0, 0, time.UTC)
	rowsBySection := buildDataInsightsRows(entries, currentMonthsThrough(beforeSeason), beforeSeason)
	rows := rowsBySection["Winter"]
	if len(rows) != 1 {
		t.Fatalf("expected one Winter row, got %d", len(rows))
	}

	row := rows[0]
	if row.Date != "December 25" {
		t.Fatalf("expected Christmas to display as December 25, got %q", row.Date)
	}
	if diff := math.Abs(row.ProjectedDollar - 100.0); diff > 1e-9 {
		t.Fatalf("expected winter projection to stay at YTD before July 1, got %.9f", row.ProjectedDollar)
	}
	if !strings.HasPrefix(row.Final, "NOT STARTED:") {
		t.Fatalf("expected Christmas to show NOT STARTED before July 1, got %q", row.Final)
	}
	if !strings.Contains(row.Final, "YoY") {
		t.Fatalf("expected NOT STARTED status to include YoY comparison, got %q", row.Final)
	}

	inSeason := time.Date(2026, time.September, 1, 12, 0, 0, 0, time.UTC)
	rowsBySection = buildDataInsightsRows(entries, currentMonthsThrough(inSeason), inSeason)
	row = rowsBySection["Winter"][0]
	expectedProjected := 100.0 * (monthsThroughSinceDate(2026, time.July, 1, time.December, 25, time.UTC) / monthsThroughSinceDate(2026, time.July, 1, time.September, 1, time.UTC))
	if diff := math.Abs(row.ProjectedDollar - expectedProjected); diff > 1e-9 {
		t.Fatalf("expected winter projection %.9f after July 1, got %.9f", expectedProjected, row.ProjectedDollar)
	}
	if !strings.HasPrefix(row.Final, "IN PROGRESS:") {
		t.Fatalf("expected Christmas to remain in progress before Dec 25, got %q", row.Final)
	}
}

// TestBuildOtherProductDataInsightsRowsSeasonalBuckets confirms other products split into
// class/occasion buckets and reuse the same holiday metadata and projection rules.
func TestBuildOtherProductDataInsightsRowsSeasonalBuckets(t *testing.T) {
	t.Parallel()

	entries := []*entry{
		{RawClassDesc: "Counter Cards", DollarSoldYTD: 999, DollarSoldPY: 888},
		{RawClassDesc: "Napkins", Occasion: "HOLIDAY", DollarSoldYTD: 75, DollarSoldPY: 70},
		{RawClassDesc: "Napkins", Occasion: "VETERAN'S DAY", DollarSoldYTD: 55, DollarSoldPY: 45},
		{RawClassDesc: "Coasters", Occasion: "HOLIDAY", DollarSoldYTD: 120, DollarSoldPY: 90},
		{RawClassDesc: "Gift Wrap", Occasion: "Mother's Day", DollarSoldYTD: 150, DollarSoldPY: 120},
		{RawClassDesc: "Alpha Everyday", DollarSoldYTD: 60, DollarSoldPY: 50},
		{RawClassDesc: "Desk Notes", DollarSoldYTD: 45, DollarSoldPY: 35},
		{RawClassDesc: "", ClassDesc: "", DollarSoldYTD: 30, DollarSoldPY: 20},
	}

	now := time.Date(2026, time.September, 1, 12, 0, 0, 0, time.UTC)
	rowsBySection := buildOtherProductDataInsightsRows(entries, currentMonthsThrough(now), now)

	if got := len(rowsBySection["Spring"]); got != 1 {
		t.Fatalf("expected one Spring row, got %d", got)
	}
	if rowsBySection["Spring"][0].Class != "Gift Wrap" || rowsBySection["Spring"][0].Occasion != "Mother's Day" {
		t.Fatalf("expected Spring row to carry class and occasion, got %+v", rowsBySection["Spring"])
	}
	if rowsBySection["Spring"][0].Date != "May 10" {
		t.Fatalf("expected Mother's Day to display May 10, got %q", rowsBySection["Spring"][0].Date)
	}
	if !strings.HasPrefix(rowsBySection["Spring"][0].Final, "COMPLETE:") {
		t.Fatalf("expected Mother's Day to be complete, got %q", rowsBySection["Spring"][0].Final)
	}

	if got := len(rowsBySection["Winter"]); got != 3 {
		t.Fatalf("expected three Winter rows, got %d", got)
	}
	if rowsBySection["Winter"][0].Class != "Coasters" || rowsBySection["Winter"][0].Occasion != "HOLIDAY" {
		t.Fatalf("expected first Winter row to be Coasters / HOLIDAY, got %+v", rowsBySection["Winter"][0])
	}
	if rowsBySection["Winter"][1].Class != "Napkins" || rowsBySection["Winter"][1].Occasion != "VETERAN'S DAY" {
		t.Fatalf("expected second Winter row to be Napkins / VETERAN'S DAY, got %+v", rowsBySection["Winter"][1])
	}
	if rowsBySection["Winter"][2].Class != "Napkins" || rowsBySection["Winter"][2].Occasion != "HOLIDAY" {
		t.Fatalf("expected third Winter row to be Napkins / HOLIDAY, got %+v", rowsBySection["Winter"][2])
	}
	if rowsBySection["Winter"][0].Date != "December 25" || rowsBySection["Winter"][1].Date != "November 11" || rowsBySection["Winter"][2].Date != "December 25" {
		t.Fatalf("expected Winter rows to use holiday dates, got %+v", rowsBySection["Winter"])
	}

	if got := len(rowsBySection["Everyday"]); got != 3 {
		t.Fatalf("expected three Everyday rows, got %d", got)
	}
	if rowsBySection["Everyday"][0].Class != "Alpha Everyday" || rowsBySection["Everyday"][1].Class != "Desk Notes" || rowsBySection["Everyday"][2].Class != "UNCLASSIFIED" {
		t.Fatalf("expected Everyday rows to be sorted alphabetically and default blanks to UNCLASSIFIED, got %+v", rowsBySection["Everyday"])
	}
}

// TestWriteDataInsightsSheetStacksOtherProductsBySeason verifies the right-hand Data
// Insights tables are still stacked Spring, Winter, then Everyday after the layout change.
func TestWriteDataInsightsSheetStacksOtherProductsBySeason(t *testing.T) {
	t.Parallel()

	f := excelize.NewFile()
	entries := []*entry{
		{RawClassDesc: "Gift Wrap", Occasion: "Mother's Day", DollarSoldYTD: 150, DollarSoldPY: 120},
		{RawClassDesc: "Napkins", Occasion: "HOLIDAY", DollarSoldYTD: 120, DollarSoldPY: 90},
		{RawClassDesc: "Napkins", Occasion: "VETERAN'S DAY", DollarSoldYTD: 55, DollarSoldPY: 45},
		{RawClassDesc: "Alpha Everyday", DollarSoldYTD: 60, DollarSoldPY: 50},
	}

	if err := writeDataInsightsSheet(f, entries); err != nil {
		t.Fatalf("writeDataInsightsSheet returned error: %v", err)
	}

	const sheetName = dataInsightsSheetName
	if got, err := f.GetCellValue(sheetName, "G5"); err != nil || got != "Spring" {
		t.Fatalf("expected G5 to contain Spring, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "G6"); err != nil || got != "Class" {
		t.Fatalf("expected G6 to contain Class, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "H6"); err != nil || got != "Occasion" {
		t.Fatalf("expected H6 to contain Occasion, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "G7"); err != nil || got != "Gift Wrap" {
		t.Fatalf("expected G7 to contain Gift Wrap, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "H7"); err != nil || got != "Mother's Day" {
		t.Fatalf("expected H7 to contain Mother's Day, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "J8"); err != nil || got != "$150.00" {
		t.Fatalf("expected J8 to contain the actual Spring YTD total, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "K8"); err != nil || got != "$120.00" {
		t.Fatalf("expected K8 to contain the actual Spring PY total, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "G10"); err != nil || got != "Winter" {
		t.Fatalf("expected G10 to contain Winter, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "G11"); err != nil || got != "Class" {
		t.Fatalf("expected G11 to contain Class, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "H11"); err != nil || got != "Occasion" {
		t.Fatalf("expected H11 to contain Occasion, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "G16"); err != nil || got != "Everyday" {
		t.Fatalf("expected G16 to contain Everyday, got %q (err=%v)", got, err)
	}
}
