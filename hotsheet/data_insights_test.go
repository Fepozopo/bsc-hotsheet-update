package hotsheet

import (
	"math"
	"strings"
	"testing"
	"time"
)

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
