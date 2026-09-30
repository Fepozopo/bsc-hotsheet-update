package hotsheet

import (
	"archive/zip"
	"io"
	"math"
	"path/filepath"
	"strconv"
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

// TestBuildMTOHistoryProfileInsufficient verifies that a partial-month launch
// or fewer than two completed months cannot produce a demand forecast.
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
		{"one completed month", []salesRecord{{Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 0, 6}}}, 1, 1},
	} {
		t.Run(tc.name, func(t *testing.T) {
			profile := buildMTOHistoryProfile(tc.records, asOf)
			if profile.usable || profile.months != tc.months || profile.years != tc.years {
				t.Errorf("%s: expected unusable profile with %d months and %d years, got %+v", tc.name, tc.months, tc.years, profile)
			}
		})
	}
}

// TestBuildMTOHistoryProfileFillsMissingMonths checks that a limited-history
// forecast uses a completed-month mean only where seasonal history is absent.
func TestBuildMTOHistoryProfileFillsMissingMonths(t *testing.T) {
	for _, tc := range []struct {
		name    string
		asOf    time.Time
		records []salesRecord
		months  int
		want    map[time.Month]float64
	}{
		{
			name:    "two completed months",
			asOf:    time.Date(2026, time.March, 25, 0, 0, 0, 0, time.UTC),
			records: []salesRecord{{Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{20, 40, 999}}},
			months:  2,
			want: map[time.Month]float64{
				time.January: 20, time.February: 40, time.March: 30, time.December: 30,
			},
		},
		{
			name:    "zero-sale completed month counts in mean",
			asOf:    time.Date(2026, time.August, 25, 0, 0, 0, 0, time.UTC),
			records: []salesRecord{{Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{0, 0, 0, 0, 20, 40, 0, 999}}},
			months:  3,
			want: map[time.Month]float64{
				time.January: 20, time.May: 20, time.June: 40, time.July: 0, time.August: 20,
			},
		},
		{
			name:    "five completed months keep their own rates",
			asOf:    time.Date(2026, time.June, 25, 0, 0, 0, 0, time.UTC),
			records: []salesRecord{{Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{10, 20, 30, 40, 50, 999}}},
			months:  5,
			want: map[time.Month]float64{
				time.January: 10, time.February: 20, time.March: 30, time.April: 40, time.May: 50, time.June: 30, time.December: 30,
			},
		},
	} {
		t.Run(tc.name, func(t *testing.T) {
			profile := buildMTOHistoryProfile(tc.records, tc.asOf)
			if !profile.usable || profile.months != tc.months || profile.years != len(tc.records) {
				t.Fatalf("%s: expected usable profile with %d months across %d years, got %+v", tc.name, tc.months, len(tc.records), profile)
			}
			for month, want := range tc.want {
				if got := profile.monthly[month-1]; math.Abs(got-want) > 1e-9 {
					t.Errorf("%s month %s: expected %v units, got %v", tc.name, month, want, got)
				}
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

// TestBuildMTORowsLimitedHistory verifies that fallback months affect both
// the projected demand and the baseline stockout for a newly launched item.
func TestBuildMTORowsLimitedHistory(t *testing.T) {
	asOf := time.Date(2026, time.September, 25, 0, 0, 0, 0, time.UTC)
	item := &inventoryEntry{SKU: "NEW", Status: "Active", OnHand: 10, SalesRecords: []salesRecord{
		{Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 20, 40, 999}},
	}}
	rows := buildMTORows([]*inventoryEntry{item}, asOf)
	if len(rows) != 1 {
		t.Fatalf("one active item: expected one MTO row, got %d", len(rows))
	}
	row := rows[0]
	if !row.hasDemand || row.state != mtoStockout || row.stockout != "Oct 2026" || row.coverage != "2 months / 1 years" || math.Abs(row.demand12-360) > 1e-9 || math.Abs(row.mto-1.0/3) > 1e-9 {
		t.Errorf("two completed months, available 10: expected 360 forecast units, Oct 2026 stockout at 1/3 month and 2 months / 1 years of coverage, got %+v", row)
	}
}

// TestMTOLimitedHistoryWorkbook checks that the saved sheet displays a
// limited-history forecast and its proposed-PO scenario uses the same fallback.
func TestMTOLimitedHistoryWorkbook(t *testing.T) {
	asOf := time.Date(2026, time.September, 25, 0, 0, 0, 0, time.UTC)
	item := &inventoryEntry{SKU: "NEW", Status: "Active", OnHand: 10, SalesRecords: []salesRecord{
		{Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{0, 0, 0, 0, 0, 0, 20, 40, 999}},
	}}
	f := newProductLineWorkbook()
	defer func() { _ = f.Close() }()
	if err := writeMTOSheet(f, []*inventoryEntry{item}, asOf); err != nil {
		t.Fatalf("write MTO for two completed months: %v", err)
	}
	path := filepath.Join(t.TempDir(), "limited-history.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save MTO for two completed months: %v", err)
	}
	written, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("reopen MTO for two completed months: %v", err)
	}
	defer func() { _ = written.Close() }()
	for cell, want := range map[string]string{"C2": "360", "D2": "Oct 2026", "I2": "2 months / 1 years"} {
		got, err := written.GetCellValue(mtoSheetName, cell, excelize.Options{RawCellValue: true})
		if err != nil || got != want {
			t.Errorf("two completed months, %s: expected %q, got %q (error %v)", cell, want, got, err)
		}
	}
	if err := written.SetCellValue(mtoSheetName, "F2", 30); err != nil {
		t.Fatalf("set proposed PO units to 30: %v", err)
	}
	if got, err := written.CalcCellValue(mtoSheetName, "H2"); err != nil || got != "Nov 2026" {
		t.Errorf("two completed months, 30 proposed units: expected Nov 2026 stockout, got %q (error %v)", got, err)
	}
	got, err := written.CalcCellValue(mtoSheetName, "G2", excelize.Options{RawCellValue: true})
	actual, parseErr := strconv.ParseFloat(got, 64)
	if err != nil || parseErr != nil || math.Abs(actual-4.0/3) > 1e-9 {
		t.Errorf("two completed months, 30 proposed units: expected MTO 4/3, got %q (calculation error %v, parse error %v)", got, err, parseErr)
	}
}

// TestMTOProposedPO verifies saved what-if formulas preserve the baseline and
// move stockout to the next seasonal peak rather than averaging across months.
func TestMTOProposedPO(t *testing.T) {
	asOf := time.Date(2026, time.September, 25, 0, 0, 0, 0, time.UTC)
	var prior, current [12]float64
	prior[0], prior[3], current[3] = 1, 120, 120
	f := newProductLineWorkbook()
	defer func() { _ = f.Close() }()
	entries := []*inventoryEntry{
		{SKU: "SPRING", Status: "Carryover", OnHand: 60, SalesRecords: []salesRecord{
			{Year: 2025, Metric: "Quantity Sold", Periods: prior},
			{Year: 2026, Metric: "Quantity Sold", Periods: current},
		}},
		{SKU: "NEW", Status: "Active", OnHand: 0},
	}
	if err := writeMTOSheet(f, entries, asOf); err != nil {
		t.Fatalf("write MTO workbook: %v", err)
	}
	path := filepath.Join(t.TempDir(), "scenario.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save MTO workbook: %v", err)
	}
	written, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("reopen MTO workbook: %v", err)
	}
	defer func() { _ = written.Close() }()
	if visible, err := written.GetSheetVisible(mtoScenarioSheetName); err != nil || visible {
		t.Fatalf("forecast helper should be hidden: visible=%v err=%v", visible, err)
	}

	for _, tc := range []struct {
		name, month, mtoText string
		input                interface{}
		mtoWant              float64
	}{
		{"blank matches baseline", "Apr 2027", "", "", 6 + 5.0/30 + (60-1.0/3)/120},
		{"next spring", "Apr 2028", "", 70, 18 + 5.0/30 + (130-120-2.0/3)/120},
		{"past horizon", "Not within 24 months", ">24", 500, 0},
		{"invalid negative", "Invalid PO units", "Invalid PO units", -1, 0},
		{"invalid fractional", "Invalid PO units", "Invalid PO units", 1.5, 0},
	} {
		t.Run(tc.name, func(t *testing.T) {
			if err := written.SetCellValue(mtoSheetName, "F3", tc.input); err != nil {
				t.Fatalf("set proposed PO %v: %v", tc.input, err)
			}
			for cell, want := range map[string]string{"H3": tc.month, "G3": tc.mtoText} {
				got, err := written.CalcCellValue(mtoSheetName, cell, excelize.Options{RawCellValue: true})
				if err != nil {
					formula, _ := written.GetCellFormula(mtoSheetName, cell)
					t.Fatalf("calculate %s for proposed PO %v (%s): %v", cell, tc.input, formula, err)
				}
				if want != "" && got != want {
					t.Errorf("proposed PO %v, %s: expected %q, got %q", tc.input, cell, want, got)
				}
				if cell == "G3" && want == "" {
					actual, parseErr := strconv.ParseFloat(got, 64)
					if parseErr != nil || math.Abs(actual-tc.mtoWant) > 1e-9 {
						t.Errorf("proposed PO %v, G3: expected MTO %v, got %q (parse error %v)", tc.input, tc.mtoWant, got, parseErr)
					}
				}
			}
			got, err := written.GetCellValue(mtoSheetName, "D3")
			if err != nil || got != "Apr 2027" {
				t.Errorf("baseline stockout after proposed PO %v: expected Apr 2027, got %q (error %v)", tc.input, got, err)
			}
		})
	}
	if err := written.SetCellValue(mtoSheetName, "F2", 5); err != nil {
		t.Fatalf("set proposed PO for missing history: %v", err)
	}
	for _, cell := range []string{"G2", "H2"} {
		if got, err := written.CalcCellValue(mtoSheetName, cell); err != nil || got != "Insufficient history" {
			t.Errorf("missing history with 5 proposed units, %s: expected Insufficient history, got %q (error %v)", cell, got, err)
		}
	}
}

// TestWriteMTOSheet checks user-visible fields, occasion-mapped seasons,
// filtering, and prompt-free PO validation in the saved workbook without
// asserting cosmetic styles or colors.
func TestWriteMTOSheet(t *testing.T) {
	asOf := time.Date(2026, time.September, 25, 0, 0, 0, 0, time.UTC)
	f := newProductLineWorkbook()
	defer func() { _ = f.Close() }()
	item := &inventoryEntry{
		SKU: "A-WM", ProductLine: "BAS", OnHand: 0, Status: "Carryover",
		ClassDesc: "Counter Cards", RawClassDesc: "Counter Cards",
		Description: "Birthday Card", Occasion: "BIRTHDAY", Foil: "Yes", CardSize: "A7",
	}
	if err := writeMTOSheet(f, []*inventoryEntry{item,
		{SKU: "C-WIN", Occasion: "CHRISTMAS", Status: "Active"},
		{SKU: "S-SPR", Occasion: "EASTER", Status: "Active"},
		{SKU: "D", Status: "Discontinued"}}, asOf); err != nil {
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
		"A1": "SKU", "B1": "Available Quantity", "C1": "Forecast Demand", "D1": "Projected Stockout Month", "E1": "MTO", "F1": "Proposed PO Units", "G1": "MTO with Proposed PO", "H1": "Stockout Month with Proposed PO", "I1": "History Coverage", "J1": "Class Description", "K1": "Season", "L1": "Occasion", "M1": "Foil", "N1": "Description", "O1": "Card Size",
		"A2": "A-WM", "B2": "0", "D2": "Sep 2026", "E2": "0", "F2": "", "I2": "0 months / 0 years", "J2": "WM - Counter Cards", "K2": "Everyday", "L2": "BIRTHDAY", "M2": "Yes", "N2": "Birthday Card", "O2": "A7",
		"A3": "C-WIN", "K3": "Winter", "L3": "CHRISTMAS",
		"A4": "S-SPR", "K4": "Spring", "L4": "EASTER",
	}
	for cell, expected := range want {
		actual, err := written.GetCellValue(mtoSheetName, cell, excelize.Options{RawCellValue: true})
		if err != nil || actual != expected {
			t.Errorf("cell %s: expected %q, got %q (error %v)", cell, expected, actual, err)
		}
	}
	if actual, err := written.GetCellValue(mtoSheetName, "A5"); err != nil || actual != "" {
		t.Errorf("A5: expected discontinued SKU omitted, got %q (error %v)", actual, err)
	}
	validations, err := written.GetDataValidations(mtoSheetName)
	if err != nil {
		t.Fatalf("cannot read saved MTO validation: %v", err)
	}
	if len(validations) != 1 {
		t.Fatalf("expected one proposed-PO validation, got %d", len(validations))
	}
	validation := validations[0]
	if validation.Sqref != "F2:F4" || validation.ShowInputMessage || validation.Prompt != nil || validation.PromptTitle != nil || !validation.ShowErrorMessage || !validation.AllowBlank || validation.Type != "whole" || validation.Operator != "between" || validation.Formula1 != "0" || validation.Formula2 != "2147483647" {
		t.Errorf("expected prompt-free F2:F4 nonnegative whole-number validation with error feedback and blanks allowed, got %+v", validation)
	}
	// Verify Excel stores the filter range through the final Card Size column.
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
		if strings.Contains(string(contents), `<autoFilter ref="$A$1:$O$1"`) {
			found = true
		}
	}
	if !found {
		t.Fatal("expected A1:O1 autofilter in saved MTO sheet")
	}
}
