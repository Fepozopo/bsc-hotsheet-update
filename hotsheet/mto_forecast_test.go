package hotsheet

import (
	"archive/zip"
	"io"
	"math"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// TestBuildMTOHistoryProfile verifies that pre-sale and partial months are not
// counted, while completed zero-sale months contribute to the seasonal profile.
func TestBuildMTOHistoryProfile(t *testing.T) {
	asOf := time.Date(2026, time.September, 25, 0, 0, 0, 0, time.UTC)
	records := []salesRecord{
		{Year: 2024, Metric: "Quantity Sold", Periods: [12]float64{0, 0, 0, 0, 0, 12, 0, 6, 8, 10, 12, 14}},
		{Year: 2025, Metric: "Quantity Sold", Periods: [12]float64{2, 4, 6, 8, 10, 24, 0, 12, 16, 20, 24, 28}},
		{Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{4, 8, 12, 16, 20, 36, 0, 18, 999}},
		{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{999}},
	}
	profile := buildMTOHistoryProfile(records, asOf)
	if !profile.usable || profile.months != 27 || profile.years != 3 {
		t.Fatalf("first sale Jun 2024, report Sep 25 2026: expected 27 completed months across 3 years and usable profile, got %+v", profile)
	}
	for _, tc := range []struct {
		month int
		want  float64
	}{
		{1, 10.0 / 3}, // Jan 2024 precedes the first sale; 2026 is weighted over 2025.
		{6, 28},       // Jun uses 2026:2025:2024 weights of 3:2:1.
		{7, 0},        // A completed zero-sales month is not dropped.
		{9, 40.0 / 3}, // The incomplete Sep 2026 value of 999 is excluded.
	} {
		if got := profile.monthly[tc.month-1]; math.Abs(got-tc.want) > 1e-9 {
			t.Errorf("month %d: expected %v units, got %v", tc.month, tc.want, got)
		}
	}
}

// TestBuildMTOHistoryProfileInsufficient verifies that a launch in a partial
// month or missing post-launch seasonal months cannot produce a full-year forecast.
func TestBuildMTOHistoryProfileInsufficient(t *testing.T) {
	asOf := time.Date(2026, time.September, 25, 0, 0, 0, 0, time.UTC)
	for _, tc := range []struct {
		name    string
		records []salesRecord
		months  int
		years   int
	}{
		{"no history", nil, 0, 0},
		{"only positive sale is in incomplete September", []salesRecord{{Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 0, 0, 7}}}, 0, 0},
		{"launched in June", []salesRecord{{Year: 2025, Metric: "Quantity Sold", Periods: [12]float64{0, 0, 0, 0, 0, 6, 0, 0, 0, 0, 0, 0}}}, 7, 1},
	} {
		t.Run(tc.name, func(t *testing.T) {
			profile := buildMTOHistoryProfile(tc.records, asOf)
			if profile.usable || profile.months != tc.months || profile.years != tc.years {
				t.Errorf("%s: expected unusable profile with %d months and %d years, got %+v", tc.name, tc.months, tc.years, profile)
			}
		})
	}
}

// TestForecastMTOWindow verifies that current and terminal months are prorated
// consistently, and stockout is located within the month that exhausts units.
func TestForecastMTOWindow(t *testing.T) {
	start := time.Date(2026, time.September, 26, 0, 0, 0, 0, time.UTC)
	end := time.Date(2026, time.November, 1, 0, 0, 0, 0, time.UTC)
	monthly := [12]float64{0, 0, 0, 0, 0, 0, 0, 0, 60, 31}
	demand, mto, stockout := forecastMTOWindow(monthly, start, end, 20)
	if math.Abs(demand-41) > 1e-9 || math.Abs(mto-(5.0/30+10.0/31)) > 1e-9 || stockout != "Oct 2026" {
		t.Errorf("Sep 26–Nov 1, available 20: expected demand=41, MTO=%v, stockout Oct 2026; got demand=%v MTO=%v stockout=%q", 5.0/30+10.0/31, demand, mto, stockout)
	}
}

// TestMonthsAfterClampsLeapDay verifies a 12-month forecast boundary does not
// roll February 29 into March in non-leap years.
func TestMonthsAfterClampsLeapDay(t *testing.T) {
	start := time.Date(2024, time.February, 29, 0, 0, 0, 0, time.UTC)
	want := time.Date(2025, time.February, 28, 0, 0, 0, 0, time.UTC)
	if got := monthsAfter(start, 12); !got.Equal(want) {
		t.Errorf("12 calendar months after %v: expected %v, got %v", start, want, got)
	}
}

// TestBuildMTORows verifies PO-inclusive availability, sold-only demand, status
// exclusions, forecast horizon, insufficient history, and stockout-first order.
func TestBuildMTORows(t *testing.T) {
	asOf := time.Date(2026, time.September, 25, 0, 0, 0, 0, time.UTC)
	history := []salesRecord{
		{Year: 2025, Metric: "Quantity Sold", Periods: [12]float64{10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10}},
		{Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{10, 10, 10, 10, 10, 10, 10, 10, 999}},
		{Year: 2026, Metric: "Quantity Returned", Periods: [12]float64{200}},
	}
	entries := []*inventoryEntry{
		{SKU: "A", Status: "Carryover", OnHand: 5, OnPO: 35, OnSO: 5, OnBO: 5, SalesRecords: history},
		{SKU: "B", Status: "Active", OnHand: 0},
		{SKU: "C", Status: "Active", OnHand: 50},
		{SKU: "D", Status: "Rundown", OnHand: 1, SalesRecords: history},
		{SKU: "E", Status: "Discontinued", OnHand: 1, SalesRecords: history},
		{SKU: "F", Status: "Carryover", OnHand: 1000, SalesRecords: history},
	}
	rows := buildMTORows(entries, asOf)
	if len(rows) != 4 {
		t.Fatalf("expected four non-rundown/non-discontinued items, got %d: %+v", len(rows), rows)
	}
	for i, want := range []string{"B", "A", "F", "C"} {
		if rows[i].item.SKU != want {
			t.Errorf("row %d: expected SKU %s, got %s", i, want, rows[i].item.SKU)
		}
	}
	if rows[0].mto != 0 || rows[0].stockout != "Sep 2026" {
		t.Errorf("out-of-stock B: expected MTO 0 and Sep 2026, got %+v", rows[0])
	}
	if rows[1].available != 30 || !rows[1].hasDemand || math.Abs(rows[1].demand12-120) > 1e-9 || rows[1].stockout != "Dec 2026" || math.Abs(rows[1].mto-3) > 1e-9 || rows[1].coverage != "20 months / 2 years" {
		t.Errorf("PO-backed SKU A: expected available 30, 120 next-year units, Dec 2026 stockout, MTO 3, 20 months / 2 years; got %+v", rows[1])
	}
	if rows[2].state != mtoBeyondHorizon || rows[3].state != mtoInsufficientHistory || rows[3].hasDemand {
		t.Errorf("expected F beyond 24 months and C with insufficient history, got F=%+v C=%+v", rows[2], rows[3])
	}
}

// TestWriteMTOSheet checks user-visible fields and filters in the saved workbook;
// it does not assert cosmetic style IDs or hardcoded colors.
func TestWriteMTOSheet(t *testing.T) {
	asOf := time.Date(2026, time.September, 25, 0, 0, 0, 0, time.UTC)
	f := newProductLineWorkbook()
	defer func() { _ = f.Close() }()
	item := &inventoryEntry{
		SKU: "A-WM", ProductLine: "BAS", OnHand: 0, Status: "Carryover",
		ClassDesc: "Counter Cards", RawClassDesc: "Counter Cards",
		Description: "Birthday Card", Occasion: "BIRTHDAY", Foil: "Yes", CardSize: "A7",
	}
	if err := writeMTOSheet(f, []*inventoryEntry{item, {SKU: "D", Status: "Discontinued"}}, asOf); err != nil {
		t.Fatalf("cannot write MTO sheet: %v", err)
	}
	path := filepath.Join(t.TempDir(), "mto.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("cannot save MTO sheet: %v", err)
	}
	written, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("cannot reopen MTO sheet: %v", err)
	}
	defer func() { _ = written.Close() }()
	want := map[string]string{
		"A1": "SKU", "B1": "Available Quantity", "C1": "Forecast Demand", "D1": "Projected Stockout Month", "E1": "MTO", "F1": "History Coverage", "G1": "Class Description", "H1": "Description", "I1": "Occasion", "J1": "Foil", "K1": "Card Size",
		"A2": "A-WM", "B2": "0", "D2": "Sep 2026", "E2": "0", "F2": "0 months / 0 years", "G2": "WM - Counter Cards", "H2": "Birthday Card", "I2": "BIRTHDAY", "J2": "Yes", "K2": "A7",
	}
	for cell, expected := range want {
		actual, err := written.GetCellValue(mtoSheetName, cell, excelize.Options{RawCellValue: true})
		if err != nil || actual != expected {
			t.Errorf("cell %s: expected %q, got %q (error %v)", cell, expected, actual, err)
		}
	}
	if actual, err := written.GetCellValue(mtoSheetName, "A3"); err != nil || actual != "" {
		t.Errorf("A3: expected discontinued SKU omitted, got %q (error %v)", actual, err)
	}
	// Verify Excel stores the filter range for all eleven requested columns.
	archive, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("cannot inspect MTO workbook: %v", err)
	}
	defer func() { _ = archive.Close() }()
	found := false
	for _, member := range archive.File {
		if !strings.HasPrefix(member.Name, "xl/worksheets/") {
			continue
		}
		file, err := member.Open()
		if err != nil {
			t.Fatalf("cannot open worksheet %s: %v", member.Name, err)
		}
		contents, err := io.ReadAll(file)
		_ = file.Close()
		if err != nil {
			t.Fatalf("cannot read worksheet %s: %v", member.Name, err)
		}
		if strings.Contains(string(contents), `<autoFilter ref="$A$1:$K$1"`) {
			found = true
		}
	}
	if !found {
		t.Fatal("expected A1:K1 autofilter in saved MTO sheet")
	}
}
