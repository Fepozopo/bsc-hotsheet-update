package hotsheet

import (
	"path/filepath"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// TestDataInsightsForecast verifies user-visible forecasts against independently
// calculated percentages, including completed-month cutoffs and missing history.
func TestDataInsightsForecast(t *testing.T) {
	t.Parallel()
	for _, tc := range []struct {
		name       string
		month      time.Month
		day        int
		hasHistory bool
		records    []salesRecord
		extra      []*inventoryEntry
		py         float64
		wantYoY    string
		wantPYTD   string
	}{
		{
			name: "short history at first day of month", month: time.March, day: 1, hasHistory: true, py: 100,
			records: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 40, 999}}},
			// 60 completed dollars + ten months at 30 = 360, versus PY 100.
			wantYoY: "+260% (limited-history estimate)", wantPYTD: "0",
		},
		{
			name: "short history at last day of month", month: time.March, day: 31, hasHistory: true, py: 100,
			records: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 40, 999}}},
			wantYoY: "+260% (limited-history estimate)", wantPYTD: "0",
		},
		{
			name: "three year weighting ignores older fourth year", month: time.January, day: 25, hasHistory: true, py: 100,
			records: []salesRecord{
				{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40}},
				{Year: 2022, Metric: "Dollars Sold", Periods: [12]float64{9000, 9000, 9000, 9000, 9000, 9000, 9000, 9000, 9000, 9000, 9000, 9000}},
				{Year: 2024, Metric: "Dollars Sold", Periods: [12]float64{20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20}},
				{Year: 2023, Metric: "Dollars Sold", Periods: [12]float64{10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10}},
			},
			// Each month is (3*40 + 2*20 + 10)/6; full year = 340.
			wantYoY: "+240%", wantPYTD: "0",
		},
		{
			name: "completed actuals plus forecast not inventory annualization", month: time.September, day: 25, hasHistory: true, py: 100,
			records: []salesRecord{
				{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10}},
				{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 20, 20, 20, 20, 20, 20, 20, 999, 999, 999, 999}},
				{Year: 2027, Metric: "Dollars Sold", Periods: [12]float64{9999}},
				{Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{9999}},
				{Year: 2026, Metric: "Quantity Issued", Periods: [12]float64{9999}},
			},
			// Jan-Aug actuals = 160; Sep-Dec estimates = 40; full PY baseline = 100, not PYTD 80.
			wantYoY: "+100%", wantPYTD: "80",
		},
		{
			name: "two comparable years use two to one weighting", month: time.January, day: 1, hasHistory: true, py: 100,
			records: []salesRecord{
				{Year: 2024, Metric: "Dollars Sold", Periods: [12]float64{10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10}},
				{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40}},
			},
			wantYoY: "+260%", wantPYTD: "0",
		},
		{
			name: "observed zero after launch contributes to mean", month: time.March, day: 25, hasHistory: true, py: 100,
			records: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 0, 999}}},
			wantYoY: "+20% (limited-history estimate)", wantPYTD: "0",
		},
		{
			name: "signed dollar corrections are not clamped to unit demand", month: time.March, day: 25, hasHistory: true, py: 100,
			records: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, -10, 999}}},
			wantYoY: "-40% (limited-history estimate)", wantPYTD: "0",
		},
		{
			name: "leading zeros do not satisfy threshold", month: time.September, day: 25, hasHistory: true, py: 100,
			records: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 0, 20, 999}}},
			wantYoY: "Insufficient history", wantPYTD: "0",
		},
		{
			name: "two SKUs in one completed month still insufficient", month: time.September, day: 25, hasHistory: true, py: 100,
			records: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 0, 20}}},
			extra:   []*inventoryEntry{{RawClassDesc: "Counter Cards", Occasion: " birthday ", SalesRecords: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 0, 40}}}}},
			wantYoY: "Insufficient history", wantPYTD: "0",
		},
		{
			name: "new SKU does not block group projection", month: time.September, day: 25, hasHistory: true, py: 100,
			records: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 20}}},
			extra:   []*inventoryEntry{{RawClassDesc: "Counter Cards", Occasion: "Birthday", SalesRecords: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 0, 40}}}}},
			// Combined July/August are 20/40; remaining four months use 30.
			wantYoY: "+80% (limited-history estimate)", wantPYTD: "0",
		},
		{
			name: "missing SKU record is not an extra zero observation", month: time.March, day: 25, hasHistory: true, py: 100,
			records: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 40}}},
			extra:   []*inventoryEntry{{RawClassDesc: "Counter Cards", Occasion: "Birthday"}},
			wantYoY: "+260% (limited-history estimate)", wantPYTD: "0",
		},
		{
			name: "absent intervening year does not supply twelve zeros", month: time.March, day: 25, hasHistory: true, py: 100,
			records: []salesRecord{
				{Year: 2024, Metric: "Dollars Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 20}},
				{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 40}},
			},
			// 60 actual + 9 months at 80/3 + December 20 = 320.
			wantYoY: "+220% (limited-history estimate)", wantPYTD: "0",
		},
		{
			name: "only incomplete month has sales", month: time.September, day: 30, hasHistory: true, py: 100,
			records: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 0, 0, 20}}},
			wantYoY: "Insufficient history", wantPYTD: "0",
		},
		{
			name: "all zero history", month: time.September, day: 25, hasHistory: true, py: 100,
			records: []salesRecord{{Year: 2025, Metric: "Dollars Sold"}},
			wantYoY: "Insufficient history", wantPYTD: "0",
		},
		{
			name: "December still projects its incomplete month", month: time.December, day: 31, hasHistory: true, py: 100,
			records: []salesRecord{
				{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 40}},
				{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 999}},
			},
			wantYoY: "+160%", wantPYTD: "110",
		},
		{
			name: "combined aligned PY uses available SKU records", month: time.March, day: 25, hasHistory: true, py: 100,
			records: []salesRecord{
				{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{1, 2}},
				{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 40}},
			},
			extra: []*inventoryEntry{{RawClassDesc: "Counter Cards", Occasion: "Birthday", SalesRecords: []salesRecord{
				{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{3, 4}},
			}}},
			// Jan/Feb PY sum to 10. Remaining months have observed PY zeros,
			// so the full-year projection is 60, not the inventory YTD of 999.
			wantYoY: "-40%", wantPYTD: "10",
		},
		{
			name: "history supplied but SKU has no records", month: time.September, day: 25, hasHistory: true, py: 100,
			wantYoY: "Insufficient history", wantPYTD: "0",
		},
		{
			name: "history omitted", month: time.September, day: 25, py: 100,
			records: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 40}}},
			wantYoY: "Insufficient history",
		},
		{
			name: "small decline rounds to unchanged without negative zero", month: time.January, day: 1, hasHistory: true, py: 100,
			records: []salesRecord{{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{8.3, 8.3, 8.3, 8.3, 8.3, 8.3, 8.3, 8.3, 8.3, 8.3, 8.3, 8.3}}},
			wantYoY: "+0%", wantPYTD: "0",
		},
		{
			name: "zero PY despite usable history", month: time.March, day: 25, hasHistory: true,
			records: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 40}}},
			wantYoY: "N/A — no PY baseline", wantPYTD: "0",
		},
		{
			name: "missing PY and history", month: time.September, day: 25,
			wantYoY: "N/A — no PY baseline",
		},
	} {
		t.Run(tc.name, func(t *testing.T) {
			t.Parallel()
			entries := append([]*inventoryEntry{{RawClassDesc: "Counter Cards", Occasion: "Birthday", DollarSoldYTD: 999, DollarSoldPY: tc.py, SalesRecords: tc.records}}, tc.extra...)
			f := excelize.NewFile()
			t.Cleanup(func() { _ = f.Close() })
			runDate := time.Date(2026, tc.month, tc.day, 0, 0, 0, 0, time.UTC)
			if err := writeDataInsightsSheet(f, entries, tc.hasHistory, runDate); err != nil {
				t.Fatalf("run %s: write outlook: %v", runDate, err)
			}
			projectionCol := "E"
			if tc.hasHistory {
				projectionCol = "F"
				assertInsightsCell(t, f, "D6", tc.wantPYTD)
				assertInsightsCell(t, f, "D7", tc.wantPYTD)
			}
			assertInsightsCell(t, f, "A6", "Birthday")
			assertInsightsCell(t, f, "A7", "Total") // All member SKUs must produce one combined row.
			assertInsightsCell(t, f, "C6", "999")
			assertInsightsCell(t, f, projectionCol+"6", tc.wantYoY)
			assertInsightsCell(t, f, projectionCol+"7", tc.wantYoY)
		})
	}
}

// TestDataInsightsHolidayLabels verifies fixed dates, stable movable labels, and
// N/A occasions in observable worksheet rows rather than seasonal metadata.
func TestDataInsightsHolidayLabels(t *testing.T) {
	t.Parallel()
	cases := []struct{ occasion, label string }{
		{"Valentine's Day", "Feb 14"}, {"ST PATRICKS DAY", "Mar 17"},
		{"Easter", "Easter"}, {"Passover", "Passover"}, {"Mother's Day", "Mother's Day"},
		{"Father's Day", "Father's Day"}, {"Independence Day", "Jul 4"},
		{"Halloween", "Oct 31"}, {"VETERANS DAY", "Nov 11"},
		{"Thanksgiving", "Thanksgiving"}, {"Hanukkah", "Hanukkah"},
		{"HOLIDAY", "Dec 25"}, {"Christmas", "Dec 25"},
		{"Graduation", "N/A"}, {"Everyday", "N/A"}, {"Birthday", "N/A"}, {"", "N/A"},
	}
	for _, tc := range cases {
		t.Run(tc.occasion, func(t *testing.T) {
			f := excelize.NewFile()
			t.Cleanup(func() { _ = f.Close() })
			if err := writeDataInsightsSheet(f, []*inventoryEntry{{RawClassDesc: "Counter Cards", Occasion: tc.occasion}}, false, time.Time{}); err != nil {
				t.Fatal(err)
			}
			assertInsightsCell(t, f, "B5", "Holiday / Date")
			assertInsightsCell(t, f, "B6", tc.label)
		})
	}
}

// TestDataInsightsSavedWorkbook verifies routing the report date/history flag
// through the workbook builder, persisted YTD / aligned PY / full-year PY column
// order, grouping, and independent card/other-product histories after saving and
// reopening the actual XLSX.
func TestDataInsightsSavedWorkbook(t *testing.T) {
	t.Parallel()
	for _, hasHistory := range []bool{false, true} {
		name := "inventory only"
		if hasHistory {
			name = "monthly history"
		}
		t.Run(name, func(t *testing.T) {
			entries := []*inventoryEntry{
				{SKU: "A", RawClassDesc: "Counter Cards", Occasion: "Birthday", DollarSoldYTD: 50, DollarSoldPY: 60, SalesRecords: []salesRecord{
					{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10}},
					{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 0, 999}},
				}},
				{SKU: "B", RawClassDesc: "Counter Cards", Occasion: " birthday ", DollarSoldYTD: 70, DollarSoldPY: 40, SalesRecords: []salesRecord{
					{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{0, 40, 999}},
				}},
				{SKU: "C", RawClassDesc: "Counter Cards Assortment", Occasion: "Christmas", DollarSoldYTD: 9, DollarSoldPY: 5},
				{SKU: "D", RawClassDesc: "Coasters", Occasion: "Easter", DollarSoldYTD: 80, DollarSoldPY: 100, SalesRecords: []salesRecord{
					{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 40, 999}},
				}},
				{SKU: "E", ClassDesc: "ZZ Desk Notes", DollarSoldYTD: 1},
				{SKU: "F", DollarSoldYTD: 2},
			}
			path, err := buildProductLineWorkbook("BAS", entries, t.TempDir(), "20260331", false, hasHistory, nil, time.Date(2026, time.March, 31, 0, 0, 0, 0, time.UTC), nil)
			if err != nil {
				t.Fatal(err)
			}
			f, err := excelize.OpenFile(path)
			if err != nil {
				t.Fatal(err)
			}
			t.Cleanup(func() { _ = f.Close() })
			assertInsightsCell(t, f, "A4", "Counter Cards")
			assertInsightsCell(t, f, "A5", "Occasion")
			assertInsightsCell(t, f, "B5", "Holiday / Date")
			assertInsightsCell(t, f, "C5", "YTD Sales")
			pyColumn := "D"
			if hasHistory {
				pyColumn = "E"
			}
			assertInsightsCell(t, f, pyColumn+"5", "PY Sales")
			assertInsightsCell(t, f, "A6", "Birthday")
			assertInsightsCell(t, f, "C6", "120")
			assertInsightsCell(t, f, pyColumn+"6", "100")
			assertInsightsCell(t, f, "A7", "Total")
			assertInsightsCell(t, f, "C7", "120")
			assertInsightsCell(t, f, pyColumn+"7", "100")
			if hasHistory {
				assertInsightsCell(t, f, "D5", "PY Sales to Date")
				assertInsightsCell(t, f, "F5", "Projected YoY")
				assertInsightsCell(t, f, "D6", "0")
				assertInsightsCell(t, f, "D7", "0")
				// Dec 2025=10, Jan=20, Feb=40: remaining nine unknown months
				// use 70/3, while December keeps 10. Projection = 280.
				assertInsightsCell(t, f, "F6", "+180% (limited-history estimate)")
				assertInsightsCell(t, f, "H4", "Coasters")
				assertInsightsCell(t, f, "H6", "Easter")
				assertInsightsCell(t, f, "I6", "Easter")
				assertInsightsCell(t, f, "K5", "PY Sales to Date")
				assertInsightsCell(t, f, "L5", "PY Sales")
				assertInsightsCell(t, f, "K6", "0")
				assertInsightsCell(t, f, "L6", "100")
				assertInsightsCell(t, f, "K7", "0")
				assertInsightsCell(t, f, "L7", "100")
				assertInsightsCell(t, f, "M6", "+260% (limited-history estimate)")
				assertInsightsCell(t, f, "H9", "Counter Cards Assortment")
				assertInsightsCell(t, f, "H14", "UNCLASSIFIED")
				assertInsightsCell(t, f, "H16", "NO OCCASION")
				assertInsightsCell(t, f, "J16", "2")
				assertInsightsCell(t, f, "H19", "ZZ Desk Notes")
			} else {
				assertInsightsCell(t, f, "E5", "Projected YoY")
				assertInsightsCell(t, f, "E6", "Insufficient history")
				assertInsightsCell(t, f, "F5", "")
				assertInsightsCell(t, f, "G4", "Coasters")
				assertInsightsCell(t, f, "K6", "Insufficient history")
				assertInsightsCell(t, f, "G9", "Counter Cards Assortment")
				assertInsightsCell(t, f, "G14", "UNCLASSIFIED")
				assertInsightsCell(t, f, "G16", "NO OCCASION")
				assertInsightsCell(t, f, "I16", "2")
				assertInsightsCell(t, f, "G19", "ZZ Desk Notes")
			}
		})
	}
}

// TestDataInsightsTotals verifies independently forecast combined histories in
// saved card and class tables, without vetoes from unprojectable occasions or
// contamination from another class. Detail forecasts must remain unchanged.
func TestDataInsightsTotals(t *testing.T) {
	t.Parallel()
	for _, tc := range []struct {
		name                           string
		month                          time.Month
		first, second                  []salesRecord
		secondPY                       float64
		wantFirst, wantSecond, wantYoY string
		wantPYTD, wantPY               string
	}{
		{
			name: "missing occasion history does not veto total", month: time.March, secondPY: 200,
			first: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 40}}},
			// The missing record supplies no observations. Forecast 360 / full PY 300.
			wantFirst: "+260% (limited-history estimate)", wantSecond: "Insufficient history",
			wantYoY: "+20% (limited-history estimate)", wantPYTD: "0", wantPY: "300",
		},
		{
			name: "no baseline occasion does not veto total", month: time.March,
			first:     []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 40}}},
			wantFirst: "+260% (limited-history estimate)", wantSecond: "N/A — no PY baseline",
			wantYoY: "+260% (limited-history estimate)", wantPYTD: "0", wantPY: "100",
		},
		{
			name: "new occasion sales participate before sufficiency checks", month: time.March, secondPY: 200,
			first:  []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 40, 999}}},
			second: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{0, 30, 999}}},
			// Combined actuals 20 + 70, then ten months at 45 = 540.
			wantFirst: "+260% (limited-history estimate)", wantSecond: "Insufficient history",
			wantYoY: "+80% (limited-history estimate)", wantPYTD: "0", wantPY: "300",
		},
		{
			name: "total can qualify when neither occasion does", month: time.September, secondPY: 200,
			first:  []salesRecord{{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 20}}},
			second: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 0, 40, 999}}},
			// Combined coverage starts Dec 2025: nine observations total 60.
			// CY actual 40 + Sep-Nov at 60/9 + December 20 = 80.
			wantFirst: "Insufficient history", wantSecond: "Insufficient history",
			wantYoY: "-73% (limited-history estimate)", wantPYTD: "0", wantPY: "300",
		},
		{
			name: "total uses its own launch and fallback rather than summing forecasts", month: time.May, secondPY: 200,
			first:  []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 40, 0, 0, 999}}},
			second: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{0, 0, 30, 50, 999}}},
			// Individual forecasts are 180 and 400. Combined Jan-Apr actuals
			// are 140, then eight months at 35 = 420, versus full PY 300.
			wantFirst: "+80% (limited-history estimate)", wantSecond: "+100% (limited-history estimate)",
			wantYoY: "+40% (limited-history estimate)", wantPYTD: "0", wantPY: "300",
		},
		{
			name: "combined coverage removes individual limited-history flag", month: time.December, secondPY: 200,
			first:  []salesRecord{{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 20}}},
			second: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 999}}},
			// Together all twelve calendar months are observed. Completed
			// actuals 110 + December 20 = 130, versus full PY 300.
			wantFirst: "Insufficient history", wantSecond: "-40% (limited-history estimate)",
			wantYoY: "-57%", wantPYTD: "0", wantPY: "300",
		},
		{
			name: "aligned PY totals retain all observed dollars", month: time.March, secondPY: 200,
			first: []salesRecord{
				{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{1, 2}},
				{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 40}},
			},
			second: []salesRecord{
				{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{3, 4}},
				{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{10, 30}},
			},
			wantFirst: "-40%", wantSecond: "-80%", wantYoY: "-67%", wantPYTD: "10", wantPY: "300",
		},
		{
			name: "two occasions in one completed month are still insufficient", month: time.March, secondPY: 200,
			first:     []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{0, 20, 999}}},
			second:    []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{0, 40, 999}}},
			wantFirst: "Insufficient history", wantSecond: "Insufficient history",
			wantYoY: "Insufficient history", wantPYTD: "0", wantPY: "300",
		},
	} {
		for _, class := range []string{"Counter Cards", "Napkins"} {
			t.Run(tc.name+"/"+class, func(t *testing.T) {
				t.Parallel()
				f := excelize.NewFile()
				t.Cleanup(func() { _ = f.Close() })
				secondClass := class
				if class == "Napkins" {
					secondClass = "NAPKINS"
				}
				entries := []*inventoryEntry{
					{RawClassDesc: class, Occasion: "Birthday", DollarSoldYTD: 50, DollarSoldPY: 100, SalesRecords: tc.first},
					{RawClassDesc: secondClass, Occasion: "Christmas", DollarSoldYTD: 70, DollarSoldPY: tc.secondPY, SalesRecords: tc.second},
					{RawClassDesc: "ZZ Unrelated", DollarSoldYTD: 9000, DollarSoldPY: 9000, SalesRecords: []salesRecord{{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{9999, 9999}}}},
					nil,
				}
				if err := writeDataInsightsSheet(f, entries, true, time.Date(2026, tc.month, 25, 0, 0, 0, 0, time.UTC)); err != nil {
					t.Fatal(err)
				}
				path := filepath.Join(t.TempDir(), "totals.xlsx")
				if err := f.SaveAs(path); err != nil {
					t.Fatal(err)
				}
				written, err := excelize.OpenFile(path)
				if err != nil {
					t.Fatal(err)
				}
				t.Cleanup(func() { _ = written.Close() })
				occasion, ytd, pytd, py, yoy := "A", "C", "D", "E", "F"
				if class == "Napkins" {
					occasion, ytd, pytd, py, yoy = "H", "J", "K", "L", "M"
				}
				assertInsightsCell(t, written, occasion+"4", class)
				assertInsightsCell(t, written, occasion+"6", "Birthday")
				assertInsightsCell(t, written, yoy+"6", tc.wantFirst)
				assertInsightsCell(t, written, occasion+"7", "Christmas")
				assertInsightsCell(t, written, yoy+"7", tc.wantSecond)
				assertInsightsCell(t, written, occasion+"8", "Total")
				assertInsightsCell(t, written, ytd+"8", "120")
				assertInsightsCell(t, written, pytd+"8", tc.wantPYTD)
				assertInsightsCell(t, written, py+"8", tc.wantPY)
				assertInsightsCell(t, written, yoy+"8", tc.wantYoY)
			})
		}
	}
}

// TestDataInsightsEmpty checks nil/empty collections and nil entries produce a
// readable empty card table and persist successfully with no invented projection.
func TestDataInsightsEmpty(t *testing.T) {
	t.Parallel()
	for _, entries := range [][]*inventoryEntry{nil, {}, {nil}} {
		f := excelize.NewFile()
		t.Cleanup(func() { _ = f.Close() })
		if err := writeDataInsightsSheet(f, entries, false, time.Time{}); err != nil {
			t.Fatalf("entries=%v: write empty outlook: %v", entries, err)
		}
		assertInsightsCell(t, f, "A6", "Total")
		assertInsightsCell(t, f, "C6", "0")
		assertInsightsCell(t, f, "E6", "N/A — no PY baseline")
		if err := f.SaveAs(filepath.Join(t.TempDir(), "empty.xlsx")); err != nil {
			t.Fatalf("entries=%v: save empty outlook: %v", entries, err)
		}
	}
}

// assertInsightsCell compares a raw workbook value with an independent expected
// string, reporting the exact user-visible cell and registering as a test helper.
func assertInsightsCell(t *testing.T, f *excelize.File, cell, want string) {
	t.Helper()
	got, err := f.GetCellValue("Data Insights", cell, excelize.Options{RawCellValue: true})
	if err != nil || got != want {
		t.Errorf("Data Insights!%s: expected %q, got %q (error %v)", cell, want, got, err)
	}
}
