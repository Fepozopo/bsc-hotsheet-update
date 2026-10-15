package hotsheet

import (
	"fmt"
	"reflect"
	"testing"
	"time"
)

// TestBuildDashboardDataEmpty verifies empty and nil inputs retain a zero All row.
func TestBuildDashboardDataEmpty(t *testing.T) {
	cases := []struct {
		name    string
		entries []*inventoryEntry
	}{
		{name: "nil"},
		{name: "empty", entries: []*inventoryEntry{}},
		{name: "nil entries", entries: []*inventoryEntry{nil, nil}},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			for _, history := range []bool{false, true} {
				got := buildDashboardData(tc.entries, history, time.Time{})
				want := []dashboardClassData{{class: "All"}}
				if !reflect.DeepEqual(got, want) {
					t.Errorf("entries=%v history=%v: got %+v, want %+v", tc.entries, history, got, want)
				}
			}
		})
	}
}

// TestBuildDashboardDataClasses verifies class precedence, alphabetical dropdown
// order, and collision-free separation of a real All class from aggregate totals.
func TestBuildDashboardDataClasses(t *testing.T) {
	cases := []struct {
		name    string
		entries []*inventoryEntry
		labels  []string
		sales   []float64
	}{
		{
			name: "raw class fallback and unclassified",
			entries: []*inventoryEntry{
				{SKU: "Z", RawClassDesc: "Zebra", ClassDesc: "Wrong", DollarSoldYTD: 1},
				{SKU: "B", ClassDesc: "Beta", DollarSoldYTD: 2},
				nil,
				{SKU: "U", DollarSoldYTD: 4},
				{SKU: "A", RawClassDesc: "Alpha", DollarSoldYTD: 8},
				{SKU: "U2", ClassDesc: "Unclassified", DollarSoldYTD: 16},
			},
			labels: []string{"All", "Alpha", "Beta", "Unclassified", "Zebra"},
			sales:  []float64{31, 8, 2, 20, 1},
		},
		{
			name: "real All class",
			entries: []*inventoryEntry{
				{SKU: "A", RawClassDesc: "All", DollarSoldYTD: 2},
				{SKU: "B", RawClassDesc: "Other", DollarSoldYTD: 3},
			},
			labels: []string{"All", "All (class)", "Other"},
			sales:  []float64{5, 2, 3},
		},
		{
			name: "synthetic labels already used by real classes",
			entries: []*inventoryEntry{
				{SKU: "A", ClassDesc: "All", DollarSoldYTD: 1},
				{SKU: "B", RawClassDesc: "All (class)", DollarSoldYTD: 2},
				{SKU: "C", RawClassDesc: "All (class 2)", DollarSoldYTD: 4},
			},
			labels: []string{"All", "All (class 2)", "All (class 3)", "All (class)"},
			sales:  []float64{7, 4, 1, 2},
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			got := buildDashboardData(tc.entries, false, time.Time{})
			if len(got) != len(tc.labels) {
				t.Fatalf("entries=%+v: got %d rows, want %d", tc.entries, len(got), len(tc.labels))
			}
			for i, row := range got {
				if row.class != tc.labels[i] || row.sales[0] != tc.sales[i] {
					t.Errorf("row %d: got label=%q sales=%v, want label=%q sales=%v", i, row.class, row.sales[0], tc.labels[i], tc.sales[i])
				}
			}
		})
	}
}

// TestBuildDashboardDataRankings verifies independent positive-only rankings,
// issue clamping, SKU tie breaks, labels, padding, all statuses, and signed sales.
// It also checks that aggregation leaves input order and entry contents unchanged.
func TestBuildDashboardDataRankings(t *testing.T) {
	entries := []*inventoryEntry{
		{SKU: "B", Description: "aaa", YTDSold: 8, YTDIssued: -50, DollarSoldYTD: 20, DollarSoldPY: 100, Status: "Rundown"},
		{SKU: "A", Description: "zzz", YTDSold: 5, YTDIssued: 3, DollarSoldYTD: -7, Status: "Carryover"},
		{SKU: "C", YTDSold: -5, YTDIssued: 2, DollarSoldYTD: 30, Status: "Inactive"},
		{SKU: "D", YTDSold: -2, YTDIssued: 5},
		{SKU: "E"},
		{SKU: "F", YTDSold: -1, DollarSoldYTD: -1},
		{SKU: "G", Description: "g", YTDSold: -8, YTDIssued: 10, DollarSoldYTD: 20, SalesRecords: []salesRecord{
			{Year: 2001, Metric: "Dollars Sold", Periods: [12]float64{999}},
		}},
	}
	order := append([]*inventoryEntry(nil), entries...)
	before := make([]inventoryEntry, len(entries))
	for i, entry := range entries {
		before[i] = *entry
		before[i].SalesRecords = append([]salesRecord(nil), entry.SalesRecords...)
	}
	got := buildDashboardData(entries, false, time.Date(2001, time.March, 15, 0, 0, 0, 0, time.UTC))
	wantShipped := [15]dashboardItem{{"A - zzz", 8}, {"B - aaa", 8}, {"D", 3}, {"G - g", 2}}
	wantDollars := [15]dashboardItem{{"C", 30}, {"B - aaa", 20}, {"G - g", 20}}
	if len(got) != 2 {
		t.Fatalf("unclassified entries: got %d rows, want All and Unclassified", len(got))
	}
	for _, row := range got {
		if row.shipped != wantShipped {
			t.Errorf("class %q: got shipped=%+v, want %+v", row.class, row.shipped, wantShipped)
		}
		if row.dollars != wantDollars {
			t.Errorf("class %q: got dollars=%+v, want %+v", row.class, row.dollars, wantDollars)
		}
		if row.sales != [2]float64{62, 100} {
			t.Errorf("class %q without history: got sales=%v, want [62 100]", row.class, row.sales)
		}
	}
	for i, entry := range entries {
		if entry != order[i] || !reflect.DeepEqual(*entry, before[i]) {
			t.Errorf("input index %d changed: got %+v, want %+v at original pointer %p", i, *entry, before[i], order[i])
		}
	}
}

// TestBuildDashboardDataTopLimit verifies each ranking keeps exactly the top 15
// when more items qualify, including SKU ties at the shipped cutoff.
func TestBuildDashboardDataTopLimit(t *testing.T) {
	var entries []*inventoryEntry
	for i := 16; i >= 0; i-- {
		entries = append(entries, &inventoryEntry{SKU: fmt.Sprintf("S%02d", i), YTDSold: 1, DollarSoldYTD: float64(i + 1)})
	}
	got := buildDashboardData(entries, false, time.Time{})
	if len(got) != 2 {
		t.Fatalf("17 unclassified entries: got %d rows, want 2", len(got))
	}
	for _, row := range got {
		if len(row.shipped) != 15 || len(row.dollars) != 15 {
			t.Fatalf("class %q: got ranking lengths %d/%d, want 15/15", row.class, len(row.shipped), len(row.dollars))
		}
		for rank := 0; rank < 15; rank++ {
			wantShipped := dashboardItem{fmt.Sprintf("S%02d", rank), 1}
			wantDollars := dashboardItem{fmt.Sprintf("S%02d", 16-rank), float64(17 - rank)}
			if row.shipped[rank] != wantShipped || row.dollars[rank] != wantDollars {
				t.Errorf("class %q rank %d: got shipped=%+v dollars=%+v, want shipped=%+v dollars=%+v", row.class, rank, row.shipped[rank], row.dollars[rank], wantShipped, wantDollars)
			}
		}
	}
}

// TestBuildDashboardDataHistory verifies report-year anchoring, completed-month
// sales cutoffs, full-year monthly sums, signed dollars, ignored metrics/years,
// duplicate records, and history taking precedence even for items without records.
func TestBuildDashboardDataHistory(t *testing.T) {
	entries := []*inventoryEntry{
		{SKU: "A", RawClassDesc: "Alpha", DollarSoldYTD: 900, DollarSoldPY: 800, SalesRecords: []salesRecord{
			{Year: 2001, Metric: "Dollars Sold", Periods: [12]float64{10, -2, 30, 11: 120}},
			{Year: 2001, Metric: "Dollars Sold", Periods: [12]float64{1, 2, 3, 11: 12}},
			{Year: 2000, Metric: "Dollars Sold", Periods: [12]float64{5, 6, 7, 11: 12}},
			{Year: 1999, Metric: "Dollars Sold", Periods: [12]float64{9999}},
			{Year: 2002, Metric: "Dollars Sold", Periods: [12]float64{9999}},
		}},
		{SKU: "B", RawClassDesc: "Beta", DollarSoldYTD: 700, DollarSoldPY: 600, SalesRecords: []salesRecord{
			{Year: 2001, Metric: "Dollars Sold", Periods: [12]float64{4, 5, 6, 11: 10}},
			{Year: 2000, Metric: "Dollars Sold", Periods: [12]float64{-1, 2, 3, 11: 4}},
		}},
		{SKU: "C", RawClassDesc: "Missing", DollarSoldYTD: 500, DollarSoldPY: 400},
	}
	for _, metric := range []string{"Quantity Sold", "Quantity Issued", "Gross Profit Percent", "Cost of Goods Sold", "Quantity Returned"} {
		for _, year := range []int{2001, 2000} {
			entries[0].SalesRecords = append(entries[0].SalesRecords, salesRecord{Year: year, Metric: metric, Periods: [12]float64{9999, 9999}})
		}
	}
	wantMonthly := [][2][12]float64{
		{{15, 5, 39, 11: 142}, {4, 8, 10, 11: 16}},
		{{11, 0, 33, 11: 132}, {5, 6, 7, 11: 12}},
		{{4, 5, 6, 11: 10}, {-1, 2, 3, 11: 4}},
		{},
	}
	cases := []struct {
		name    string
		month   time.Month
		history bool
		sales   [][2]float64
	}{
		{"January has no completed months", time.January, true, [][2]float64{{}, {}, {}, {}}},
		{"February includes only January", time.February, true, [][2]float64{{15, 4}, {11, 5}, {4, -1}, {}}},
		{"March excludes March for both years", time.March, true, [][2]float64{{20, 12}, {11, 11}, {9, 1}, {}}},
		{"December excludes December for both years", time.December, true, [][2]float64{{59, 22}, {44, 18}, {15, 4}, {}}},
		{"without history uses inventory totals", time.March, false, [][2]float64{{2100, 1800}, {900, 800}, {700, 600}, {500, 400}}},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			runDate := time.Date(2001, tc.month, 15, 0, 0, 0, 0, time.UTC)
			got := buildDashboardData(entries, tc.history, runDate)
			if len(got) != len(tc.sales) {
				t.Fatalf("date=%v history=%v: got %d rows, want %d", runDate, tc.history, len(got), len(tc.sales))
			}
			for i, row := range got {
				if row.sales != tc.sales[i] {
					t.Errorf("date=%v history=%v class=%q: got sales=%v, want %v", runDate, tc.history, row.class, row.sales, tc.sales[i])
				}
				if row.monthly != wantMonthly[i] {
					t.Errorf("date=%v history=%v class=%q: got monthly=%v, want %v", runDate, tc.history, row.class, row.monthly, wantMonthly[i])
				}
				if row.shipped != [15]dashboardItem{} {
					t.Errorf("date=%v history=%v class=%q: got shipped=%+v, want zero inventory shipments", runDate, tc.history, row.class, row.shipped)
				}
				wantTop := dashboardItem{"A", 900}
				if i == 2 {
					wantTop = dashboardItem{"B", 700}
				} else if i == 3 {
					wantTop = dashboardItem{"C", 500}
				}
				if row.dollars[0] != wantTop {
					t.Errorf("date=%v history=%v class=%q: got top dollars=%+v, want inventory value %+v", runDate, tc.history, row.class, row.dollars[0], wantTop)
				}
			}
		})
	}
}
