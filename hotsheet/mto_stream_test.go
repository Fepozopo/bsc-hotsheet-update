package hotsheet

import (
	"fmt"
	"math"
	"strconv"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// TestMTOStreamedScenarios verifies saved helper ranges remain aligned with
// multiple visible SKUs across whole-month and leap-year partial-month horizons.
// A SKU without usable demand must not consume another SKU's forecast segments.
func TestMTOStreamedScenarios(t *testing.T) {
	for _, tc := range []struct {
		name                  string
		asOf                  time.Time
		segments              int
		firstTotal, lastTotal float64
	}{
		{
			name:     "whole months across year boundary",
			asOf:     time.Date(2023, time.December, 31, 0, 0, 0, 0, time.UTC),
			segments: 24, firstTotal: 29, lastTotal: 696,
		},
		{
			name:     "partial leap February through non-leap February",
			asOf:     time.Date(2024, time.February, 27, 0, 0, 0, 0, time.UTC),
			segments: 25, firstTotal: 2, lastTotal: 2 + 23*29 + 29*27.0/28,
		},
	} {
		t.Run(tc.name, func(t *testing.T) {
			t.Parallel()
			entries := []*inventoryEntry{
				{SKU: "BETA", OnHand: 116, SalesRecords: []salesRecord{
					{Year: 2023, Metric: "Quantity Sold", Periods: [12]float64{58, 58, 58, 58, 58, 58, 58, 58, 58, 58, 58, 58}},
				}},
				{SKU: "NO-HISTORY"},
				{SKU: "ALPHA", OnHand: 29, SalesRecords: []salesRecord{
					{Year: 2023, Metric: "Quantity Sold", Periods: [12]float64{29, 29, 29, 29, 29, 29, 29, 29, 29, 29, 29, 29}},
				}},
			}
			path, err := buildProductLineWorkbook("BAS", entries, t.TempDir(), "20240101", false, true, nil, tc.asOf, nil)
			if err != nil {
				t.Fatalf("build workbook as of %s: %v", tc.asOf, err)
			}
			written, err := excelize.OpenFile(path)
			if err != nil {
				t.Fatalf("open saved workbook %s: %v", path, err)
			}
			t.Cleanup(func() { _ = written.Close() })
			if visible, err := written.GetSheetVisible("MTO Forecast Data"); err != nil || visible {
				t.Fatalf("expected saved forecast helper to be hidden, got visible=%v error=%v", visible, err)
			}
			helper, err := written.GetRows("MTO Forecast Data", excelize.Options{RawCellValue: true})
			if err != nil {
				t.Fatalf("read saved helper data: %v", err)
			}
			if want := 2*tc.segments + 1; len(helper) != want {
				t.Fatalf("as of %s: expected blank row and %d segments per SKU (%d rows), got %d rows", tc.asOf, tc.segments, want, len(helper))
			}
			for _, check := range []struct {
				row      int
				sku      string
				proposed int
				baseline float64
			}{
				{3, "ALPHA", 29, 1},
				{4, "BETA", 58, 2},
			} {
				if got, err := written.GetCellValue("MTO", fmt.Sprintf("A%d", check.row)); err != nil || got != check.sku {
					t.Fatalf("MTO row %d: expected SKU %s, got %q (error %v)", check.row, check.sku, got, err)
				}
				if err := written.SetCellValue("MTO", fmt.Sprintf("I%d", check.row), check.proposed); err != nil {
					t.Fatalf("SKU %s: set proposed PO to %d: %v", check.sku, check.proposed, err)
				}
				// Excelize evaluates the saved formula against the streamed helper
				// values; no cached result or live Excel process is required.
				got, err := written.CalcCellValue("MTO", fmt.Sprintf("J%d", check.row), excelize.Options{RawCellValue: true})
				actual, parseErr := strconv.ParseFloat(got, 64)
				if want := check.baseline + 1; err != nil || parseErr != nil || math.Abs(actual-want) > 1e-9 {
					t.Errorf("SKU %s with %d proposed units: expected MTO %v, got %q (calculation error %v, parse error %v)", check.sku, check.proposed, want, got, err, parseErr)
				}
			}
			for _, check := range []struct {
				cell string
				want float64
			}{
				{"D2", tc.firstTotal},
				{fmt.Sprintf("D%d", tc.segments+1), tc.lastTotal},
				{fmt.Sprintf("C%d", tc.segments+2), 0},
				{fmt.Sprintf("E%d", tc.segments+2), 0},
				{fmt.Sprintf("D%d", 2*tc.segments+1), 2 * tc.lastTotal},
			} {
				got, err := written.GetCellValue("MTO Forecast Data", check.cell, excelize.Options{RawCellValue: true})
				actual, parseErr := strconv.ParseFloat(got, 64)
				if err != nil || parseErr != nil || math.Abs(actual-check.want) > 1e-9 {
					t.Errorf("as of %s, helper %s: expected %v, got %q (read error %v, parse error %v)", tc.asOf, check.cell, check.want, got, err, parseErr)
				}
			}
		})
	}
}
