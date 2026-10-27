package hotsheet

import (
	"path/filepath"
	"strconv"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// TestMTOStockoutDates verifies saved baseline dates and live proposed-PO dates
// against independently specified calendar dates, including partial months,
// leap days, midnight boundaries, seasonal gaps, and non-date forecast states.
func TestMTOStockoutDates(t *testing.T) {
	regular := [12]float64{31, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31}
	for _, tc := range []struct {
		name                string
		asOf                string
		monthly             [12]float64
		available           int
		proposed            interface{}
		baseline, projected string
		noHistory           bool
	}{
		{
			name: "daily estimate and added PO", asOf: "2026-12-31", monthly: regular,
			available: 10, proposed: 10, baseline: "2027-01-11", projected: "2027-01-21",
		},
		{
			name: "blank PO matches baseline", asOf: "2026-12-31", monthly: regular,
			available: 10, proposed: "", baseline: "2027-01-11", projected: "2027-01-11",
		},
		{
			name: "within-day depletion", asOf: "2026-12-31",
			monthly:   [12]float64{62, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31},
			available: 1, proposed: 1, baseline: "2027-01-01", projected: "2027-01-02",
		},
		{
			name: "partial month uses remaining days", asOf: "2026-09-25",
			monthly:   [12]float64{31, 31, 31, 31, 31, 31, 31, 31, 60, 31, 31, 31},
			available: 20, proposed: 10, baseline: "2026-10-11", projected: "2026-10-21",
		},
		{
			name: "year boundary", asOf: "2026-12-30", monthly: regular,
			available: 1, proposed: 1, baseline: "2027-01-01", projected: "2027-01-02",
		},
		{
			name: "leap February", asOf: "2024-02-27",
			monthly:   [12]float64{31, 29, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31},
			available: 1, proposed: 1, baseline: "2024-02-29", projected: "2024-03-01",
		},
		{
			name: "non-leap February", asOf: "2025-02-27",
			monthly:   [12]float64{31, 28, 31, 31, 31, 31, 31, 31, 31, 31, 31, 31},
			available: 1, proposed: 1, baseline: "2025-03-01", projected: "2025-03-02",
		},
		{
			name: "PO skips zero-demand months to next seasonal peak", asOf: "2026-09-25",
			monthly:   [12]float64{1, 0, 0, 120},
			available: 60, proposed: 70, baseline: "2027-04-15", projected: "2028-04-03",
		},
		{
			name: "exact horizon and one unit beyond", asOf: "2025-12-31", monthly: regular,
			available: 744, proposed: 1, baseline: "2028-01-01", projected: "Not within 24 months",
		},
		{
			name: "baseline beyond horizon", asOf: "2025-12-31", monthly: regular,
			available: 745, proposed: 0, baseline: "Not within 24 months", projected: "Not within 24 months",
		},
		{
			name: "no history", asOf: "2026-12-31", noHistory: true,
			available: 10, proposed: 10, baseline: "Insufficient history", projected: "Insufficient history",
		},
		{
			name: "zero shipments", asOf: "2026-12-31",
			available: 10, proposed: 10, baseline: "Insufficient history", projected: "Insufficient history",
		},
		{
			name: "already depleted without history", asOf: "2026-12-31", noHistory: true,
			available: 0, proposed: 5, baseline: "2026-12-31", projected: "Insufficient history",
		},
		{
			name: "PO does not cover negative availability", asOf: "2026-12-31", monthly: regular,
			available: -5, proposed: 3, baseline: "2026-12-31", projected: "2026-12-31",
		},
		{
			name: "negative PO", asOf: "2026-12-31", monthly: regular,
			available: 10, proposed: -1, baseline: "2027-01-11", projected: "Invalid PO units",
		},
		{
			name: "fractional PO", asOf: "2026-12-31", monthly: regular,
			available: 10, proposed: 1.5, baseline: "2027-01-11", projected: "Invalid PO units",
		},
		{
			name: "nonnumeric PO", asOf: "2026-12-31", monthly: regular,
			available: 10, proposed: "invalid", baseline: "2027-01-11", projected: "Invalid PO units",
		},
	} {
		t.Run(tc.name, func(t *testing.T) {
			t.Parallel()
			asOf, err := time.Parse("2006-01-02", tc.asOf)
			if err != nil {
				t.Fatalf("parse report date %q: %v", tc.asOf, err)
			}
			item := &inventoryEntry{SKU: "TEST", OnHand: tc.available}
			if !tc.noHistory {
				item.SalesRecords = []salesRecord{{Year: asOf.Year() - 1, Metric: "Quantity Sold", Periods: tc.monthly}}
			}
			f := newProductLineWorkbook()
			t.Cleanup(func() { _ = f.Close() })
			writeMTOTestSheets(t, f, []*inventoryEntry{item}, asOf)
			path := filepath.Join(t.TempDir(), "stockout.xlsx")
			if err := f.SaveAs(path); err != nil {
				t.Fatalf("save stockout dates as of %s: %v", tc.asOf, err)
			}
			written, err := excelize.OpenFile(path)
			if err != nil {
				t.Fatalf("reopen stockout dates as of %s: %v", tc.asOf, err)
			}
			t.Cleanup(func() { _ = written.Close() })

			assertMTOStockoutDate(t, written, "H2", tc.baseline, false)
			assertMTOStockoutDate(t, written, "K2", tc.baseline, true)
			if err := written.SetCellValue("MTO", "I2", tc.proposed); err != nil {
				t.Fatalf("set proposed units %v: %v", tc.proposed, err)
			}
			assertMTOStockoutDate(t, written, "K2", tc.projected, true)
			// A what-if change must never alter the baseline depletion date.
			assertMTOStockoutDate(t, written, "H2", tc.baseline, false)
		})
	}
}

// assertMTOStockoutDate checks cell's saved value or calculated formula in f
// against want (an ISO date or forecast status). Dates must be numeric Excel
// dates, not text; saved dates must also display the full month, day, and year.
// Failures are reported through t with the cell and expected value.
func assertMTOStockoutDate(t *testing.T, f *excelize.File, cell, want string, calculate bool) {
	t.Helper()
	var got string
	var err error
	if calculate {
		// Excelize evaluates formulas against the saved helper data without
		// depending on Excel or a cached calculation result.
		got, err = f.CalcCellValue("MTO", cell, excelize.Options{RawCellValue: true})
	} else {
		got, err = f.GetCellValue("MTO", cell, excelize.Options{RawCellValue: true})
	}
	if err != nil {
		t.Fatalf("read %s (calculate=%v), expected %q: %v", cell, calculate, want, err)
	}
	date, dateErr := time.Parse("2006-01-02", want)
	if dateErr != nil {
		if got != want {
			t.Errorf("%s: expected status %q, got %q", cell, want, got)
		}
		return
	}
	serial, err := strconv.ParseFloat(got, 64)
	if err != nil {
		t.Fatalf("%s: expected numeric Excel date for %s, got %q: %v", cell, want, got, err)
	}
	// Excelize decodes the workbook's standard 1900-system date serial; the
	// expected date above is independent of the forecasting implementation.
	actual, err := excelize.ExcelDateToTime(serial, false)
	if err != nil || !actual.Equal(date) {
		t.Errorf("%s: expected date %s, got %s (serial %q, error %v)", cell, want, actual, got, err)
	}
	if !calculate {
		formatted, err := f.GetCellValue("MTO", cell)
		if expected := date.Format("Jan 2, 2006"); err != nil || formatted != expected {
			t.Errorf("%s: expected displayed date %q, got %q (error %v)", cell, expected, formatted, err)
		}
	}
}
