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

	entries := []*inventoryEntry{{
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
	if !strings.HasPrefix(row.YoYDisplay, "IN PROGRESS:") {
		t.Fatalf("expected in-progress status before year end, got %q", row.YoYDisplay)
	}

	rowsBySection = buildDataInsightsRows(entries, currentMonthsThrough(time.Date(2026, time.December, 31, 12, 0, 0, 0, time.UTC)), time.Date(2026, time.December, 31, 12, 0, 0, 0, time.UTC))
	row = rowsBySection["Spring"][0]
	if diff := math.Abs(row.ProjectedDollar - 100.0); diff > 1e-9 {
		t.Fatalf("expected projected sales to match YTD at the end of the season, got %.9f", row.ProjectedDollar)
	}
	if !strings.HasPrefix(row.YoYDisplay, "IN PROGRESS:") {
		t.Fatalf("expected status to remain in progress on Dec 31, got %q", row.YoYDisplay)
	}
}

// TestGraduationProjectionWindow confirms a normal spring holiday uses its holiday date
// and flips to complete after the event day.
func TestGraduationProjectionWindow(t *testing.T) {
	t.Parallel()

	entries := []*inventoryEntry{{
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
	if !strings.HasPrefix(row.YoYDisplay, "IN PROGRESS:") {
		t.Fatalf("expected Graduation to be in progress before June 15, got %q", row.YoYDisplay)
	}

	afterCutoff := time.Date(2026, time.June, 16, 12, 0, 0, 0, time.UTC)
	rowsBySection = buildDataInsightsRows(entries, currentMonthsThrough(afterCutoff), afterCutoff)
	row = rowsBySection["Spring"][0]
	if row.Date != "mid-June" {
		t.Fatalf("expected Graduation to still display as mid-June, got %q", row.Date)
	}
	if !strings.HasPrefix(row.YoYDisplay, "COMPLETE:") {
		t.Fatalf("expected Graduation to be complete after June 15, got %q", row.YoYDisplay)
	}
}

// TestBuildDataInsightsRowsWinterProjectionStartsJune15 confirms winter holidays stay at
// YTD before July 1 and switch to in-season projection afterward.
func TestBuildDataInsightsRowsWinterProjectionStartsJune15(t *testing.T) {
	t.Parallel()

	entries := []*inventoryEntry{{
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
	if !strings.HasPrefix(row.YoYDisplay, "NOT STARTED:") {
		t.Fatalf("expected Christmas to show NOT STARTED before July 1, got %q", row.YoYDisplay)
	}
	if !strings.Contains(row.YoYDisplay, "YoY") {
		t.Fatalf("expected NOT STARTED status to include YoY comparison, got %q", row.YoYDisplay)
	}

	inSeason := time.Date(2026, time.September, 1, 12, 0, 0, 0, time.UTC)
	rowsBySection = buildDataInsightsRows(entries, currentMonthsThrough(inSeason), inSeason)
	row = rowsBySection["Winter"][0]
	expectedProjected := 100.0 * (monthsThroughSinceDate(2026, time.June, 15, time.December, 25, time.UTC) / monthsThroughSinceDate(2026, time.June, 15, time.September, 1, time.UTC))
	if diff := math.Abs(row.ProjectedDollar - expectedProjected); diff > 1e-9 {
		t.Fatalf("expected winter projection %.9f after June 15, got %.9f", expectedProjected, row.ProjectedDollar)
	}
	if !strings.HasPrefix(row.YoYDisplay, "IN PROGRESS:") {
		t.Fatalf("expected Christmas to remain in progress before Dec 25, got %q", row.YoYDisplay)
	}
}

// TestBuildOtherProductsDataInsightsRowsGroupedByClass confirms other products now group into
// one class bucket per table while still reusing the same holiday metadata and projection rules.
func TestBuildOtherProductsDataInsightsRowsGroupedByClass(t *testing.T) {
	t.Parallel()

	entries := []*inventoryEntry{
		{RawClassDesc: "Counter Cards", DollarSoldYTD: 999, DollarSoldPY: 888},
		{RawClassDesc: "Napkins", Occasion: "HOLIDAY", DollarSoldYTD: 75, DollarSoldPY: 70},
		{RawClassDesc: "Napkins", Occasion: "VETERAN'S DAY", DollarSoldYTD: 55, DollarSoldPY: 45},
		{RawClassDesc: "Coasters", Occasion: "HOLIDAY", DollarSoldYTD: 120, DollarSoldPY: 90},
		{RawClassDesc: "Gift Wrap", Occasion: "INDEPENDENCE DAY", DollarSoldYTD: 150, DollarSoldPY: 120},
		{RawClassDesc: "Alpha Everyday", DollarSoldYTD: 60, DollarSoldPY: 50},
		{RawClassDesc: "Desk Notes", DollarSoldYTD: 45, DollarSoldPY: 35},
		{RawClassDesc: "", ClassDesc: "", DollarSoldYTD: 30, DollarSoldPY: 20},
	}

	now := time.Date(2026, time.September, 1, 12, 0, 0, 0, time.UTC)
	rowsByClass := buildOtherProductsDataInsightsRows(entries, currentMonthsThrough(now), now)

	if got := len(rowsByClass); got != 6 {
		t.Fatalf("expected six class buckets, got %d", got)
	}
	if got := strings.Join(sortedOtherProductsDataInsightsClasses(rowsByClass), "|"); got != "Alpha Everyday|Coasters|Desk Notes|Gift Wrap|Napkins|UNCLASSIFIED" {
		t.Fatalf("unexpected class order: %q", got)
	}

	giftWrapRows := rowsByClass["Gift Wrap"]
	if len(giftWrapRows) != 1 {
		t.Fatalf("expected one Gift Wrap row, got %d", len(giftWrapRows))
	}
	if giftWrapRows[0].Occasion != "INDEPENDENCE DAY" || giftWrapRows[0].Date != "July 4" {
		t.Fatalf("expected Gift Wrap to keep Independence Day metadata, got %+v", giftWrapRows[0])
	}
	if !strings.HasPrefix(giftWrapRows[0].YoYDisplay, "COMPLETE:") {
		t.Fatalf("expected Independence Day to be complete, got %q", giftWrapRows[0].YoYDisplay)
	}

	napkinsRows := rowsByClass["Napkins"]
	if got := len(napkinsRows); got != 2 {
		t.Fatalf("expected two Napkins rows, got %d", got)
	}
	if napkinsRows[0].Occasion != "VETERAN'S DAY" || napkinsRows[0].Date != "November 11" {
		t.Fatalf("expected Napkins to sort Veterans Day before Holiday, got %+v", napkinsRows[0])
	}
	if napkinsRows[1].Occasion != "HOLIDAY" || napkinsRows[1].Date != "December 25" {
		t.Fatalf("expected Napkins Holiday row to keep its winter date, got %+v", napkinsRows[1])
	}

	alphaEverydayRows := rowsByClass["Alpha Everyday"]
	if got := len(alphaEverydayRows); got != 1 {
		t.Fatalf("expected one Alpha Everyday row, got %d", got)
	}
	if alphaEverydayRows[0].Occasion != "NO OCCASION" || alphaEverydayRows[0].Date != "N/A" {
		t.Fatalf("expected Alpha Everyday to remain an everyday row, got %+v", alphaEverydayRows[0])
	}
	if alphaEverydayRows[0].Section != "Everyday" {
		t.Fatalf("expected Alpha Everyday to preserve its Everyday section, got %q", alphaEverydayRows[0].Section)
	}

	unclassifiedRows := rowsByClass["UNCLASSIFIED"]
	if got := len(unclassifiedRows); got != 1 {
		t.Fatalf("expected one UNCLASSIFIED row, got %d", got)
	}
	if unclassifiedRows[0].Class != "UNCLASSIFIED" {
		t.Fatalf("expected blank classes to normalize to UNCLASSIFIED, got %+v", unclassifiedRows[0])
	}
}

// TestOtherProductsClassTotalYoYDisplayMixedClassUsesProjectedYoY confirms mixed class tables
// avoid seasonal status wording on the total row because the aggregate includes everyday sales.
func TestOtherProductsClassTotalYoYDisplayMixedClassUsesProjectedYoY(t *testing.T) {
	t.Parallel()

	rows := []dataInsightsRow{
		{Section: "Spring", YoYDisplay: "IN PROGRESS: 25.00% YoY"},
		{Section: "Everyday", YoYDisplay: "10.00%"},
	}

	got := otherProductsClassTotalYoYDisplay(150, 100, rows)
	want := formatYoYFromProjectedSales(150, 100)
	if got != want {
		t.Fatalf("expected mixed class total display %q, got %q", want, got)
	}
}

// TestWriteDataInsightsSheetGroupsOtherProductsByClass verifies the right-hand Data Insights
// tables are stacked by class name, with the class shown in the table title instead of a column.
func TestWriteDataInsightsSheetGroupsOtherProductsByClass(t *testing.T) {
	t.Parallel()

	f := excelize.NewFile()
	entries := []*inventoryEntry{
		{RawClassDesc: "Gift Wrap", Occasion: "Mother's Day", DollarSoldYTD: 150, DollarSoldPY: 120},
		{RawClassDesc: "Napkins", Occasion: "HOLIDAY", DollarSoldYTD: 120, DollarSoldPY: 90},
		{RawClassDesc: "Napkins", Occasion: "VETERAN'S DAY", DollarSoldYTD: 55, DollarSoldPY: 45},
		{RawClassDesc: "Alpha Everyday", DollarSoldYTD: 60, DollarSoldPY: 50},
	}

	if err := writeDataInsightsSheet(f, entries); err != nil {
		t.Fatalf("writeDataInsightsSheet returned error: %v", err)
	}

	const sheetName = dataInsightsSheetName
	if got, err := f.GetCellValue(sheetName, "G5"); err != nil || got != "Alpha Everyday" {
		t.Fatalf("expected G5 to contain the first class title, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "G6"); err != nil || got != "Occasion" {
		t.Fatalf("expected G6 to contain Occasion, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "H6"); err != nil || got != "Date" {
		t.Fatalf("expected H6 to contain Date, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "G7"); err != nil || got != "NO OCCASION" {
		t.Fatalf("expected G7 to contain the Alpha Everyday occasion, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "I8"); err != nil || got != "$60.00" {
		t.Fatalf("expected I8 to contain the Alpha Everyday YTD total, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "J8"); err != nil || got != "$50.00" {
		t.Fatalf("expected J8 to contain the Alpha Everyday PY total, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "G10"); err != nil || got != "Gift Wrap" {
		t.Fatalf("expected G10 to contain the second class title, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "G11"); err != nil || got != "Occasion" {
		t.Fatalf("expected G11 to contain Occasion, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "G15"); err != nil || got != "Napkins" {
		t.Fatalf("expected G15 to contain the third class title, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "G17"); err != nil || got != "VETERAN'S DAY" {
		t.Fatalf("expected G17 to contain the first Napkins occasion, got %q (err=%v)", got, err)
	}
	if got, err := f.GetCellValue(sheetName, "G18"); err != nil || got != "HOLIDAY" {
		t.Fatalf("expected G18 to contain the second Napkins occasion, got %q (err=%v)", got, err)
	}
}
