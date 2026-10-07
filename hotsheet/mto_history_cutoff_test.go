package hotsheet

import (
	"math"
	"strconv"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// TestMTOKeyAccountHistoryCutoff verifies the workbook's baseline and proposed-PO
// forecasts use only 2026+ sales and issues for every designated 2021 SKU, while
// unlisted SKUs and other product lines retain their historical forecasts.
func TestMTOKeyAccountHistoryCutoff(t *testing.T) {
	history := []salesRecord{
		{Year: 2025, Metric: "Quantity Sold", Periods: [12]float64{90, 90, 90, 90, 90, 90, 90, 90, 90, 90, 90, 90}},
		{Year: 2025, Metric: "Quantity Issued", Periods: [12]float64{10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10}},
		{Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{10, 20, 999}},
		{Year: 2026, Metric: "Quantity Issued", Periods: [12]float64{5, 5, 999}},
	}
	// forecastCase defines the expected user-visible forecast for one SKU and date.
	type forecastCase struct {
		name, sku, productLine, coverage, stockout, proposedStockout string
		asOf                                                         time.Time
		records                                                      []salesRecord
		demand, mto, proposedMTO                                     float64
	}
	cases := []forecastCase{
		{
			name: "unlisted SKU", sku: "BD1001", productLine: "2021",
			coverage: "14 months / 2 years", stockout: "Mar 2026", proposedStockout: "Mar 2026",
			demand: 3280.0 / 3, mto: 0.3, proposedMTO: 0.6,
		},
		{
			name: "listed SKU in other product line", sku: "BD1001FJ", productLine: "BAS",
			coverage: "14 months / 2 years", stockout: "Mar 2026", proposedStockout: "Mar 2026",
			demand: 3280.0 / 3, mto: 0.3, proposedMTO: 0.6,
		},
		{
			name: "listed SKU without product line", sku: "BD1001FJ",
			coverage: "14 months / 2 years", stockout: "Mar 2026", proposedStockout: "Mar 2026",
			demand: 3280.0 / 3, mto: 0.3, proposedMTO: 0.6,
		},
		{
			name: "later years remain eligible", sku: "BD1001FJ", productLine: "2021",
			asOf: time.Date(2027, time.March, 1, 0, 0, 0, 0, time.UTC),
			records: []salesRecord{
				{Year: 2025, Metric: "Quantity Sold", Periods: [12]float64{999}},
				{Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20}},
				{Year: 2027, Metric: "Quantity Sold", Periods: [12]float64{40, 40, 999}},
			},
			coverage: "14 months / 2 years", stockout: "Apr 2027", proposedStockout: "Jun 2027",
			demand: 800.0 / 3, mto: 1.5, proposedMTO: 3,
		},
	}
	// Keep the expected SKU list independent of the production eligibility rule.
	for _, sku := range []string{
		"BD1001FJ", "BD1006F", "BD1028", "BD1044", "BD1046F", "BD1048FJ",
		"BP1002", "CS1003", "CS1005", "FC1005", "GR1016FJ", "HY1048FB",
		"HY1071", "HY1071B", "LV1010", "MD1011", "MI1011", "MI1014",
		"PJ1011", "PJ1031", "PJ1032", "RL1003", "SP1001F", "TY1007",
		"TY1019", "TY1019B", "TY1020", "TY1023B", "TY1025", "TY1025B",
		"TY1028F", "VD1029J",
	} {
		cases = append(cases, forecastCase{
			name: sku, sku: sku, productLine: "2021",
			coverage: "2 months / 1 years", stockout: "Apr 2026", proposedStockout: "Jun 2026",
			demand: 240, mto: 1.5, proposedMTO: 3,
		})
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			asOf := tc.asOf
			if asOf.IsZero() {
				// The first day of March makes February a completed month.
				asOf = time.Date(2026, time.March, 1, 0, 0, 0, 0, time.UTC)
			}
			records := tc.records
			if records == nil {
				records = history
			}
			item := &inventoryEntry{SKU: tc.sku, ProductLine: tc.productLine, OnHand: 30, SalesRecords: records}
			f := newProductLineWorkbook()
			t.Cleanup(func() { _ = f.Close() })
			if err := writeMTOSheet(f, []*inventoryEntry{item}, asOf); err != nil {
				t.Fatalf("write MTO for SKU %s, product line %s, as of %s: %v", tc.sku, tc.productLine, asOf, err)
			}
			if err := f.SetCellValue("MTO", "F2", 30); err != nil {
				t.Fatalf("set proposed PO to 30: %v", err)
			}
			for _, check := range []struct{ cell, want string }{
				{"I2", tc.coverage}, {"D2", tc.stockout},
			} {
				if got, err := f.GetCellValue("MTO", check.cell); err != nil || got != check.want {
					t.Errorf("%s: expected %q, got %q (error %v)", check.cell, check.want, got, err)
				}
			}
			if got, err := f.CalcCellValue("MTO", "H2"); err != nil || got != tc.proposedStockout {
				t.Errorf("30 proposed units: expected stockout %q, got %q (error %v)", tc.proposedStockout, got, err)
			}
			for _, check := range []struct {
				cell string
				want float64
			}{
				{"C2", tc.demand}, {"E2", tc.mto}, {"G2", tc.proposedMTO},
			} {
				var got string
				var err error
				if check.cell == "G2" {
					// Excelize evaluates the live proposed-PO formula rather than
					// reading a cached value, which has not been calculated by Excel.
					got, err = f.CalcCellValue("MTO", check.cell, excelize.Options{RawCellValue: true})
				} else {
					got, err = f.GetCellValue("MTO", check.cell, excelize.Options{RawCellValue: true})
				}
				actual, parseErr := strconv.ParseFloat(got, 64)
				if err != nil || parseErr != nil || math.Abs(actual-check.want) > 1e-9 {
					t.Errorf("%s: expected %v, got %q (read error %v, parse error %v)", check.cell, check.want, got, err, parseErr)
				}
			}
		})
	}
}

// TestMTOKeyAccountInsufficientHistory verifies excluded years cannot supply
// forecast coverage when fewer than two eligible completed months remain.
func TestMTOKeyAccountInsufficientHistory(t *testing.T) {
	prior := salesRecord{Year: 2025, Metric: "Quantity Sold", Periods: [12]float64{100, 100}}
	for _, tc := range []struct {
		name     string
		records  []salesRecord
		asOf     time.Time
		coverage string
	}{
		{"nil history", nil, time.Date(2026, time.March, 1, 0, 0, 0, 0, time.UTC), "0 months / 0 years"},
		{"empty history", []salesRecord{}, time.Date(2026, time.March, 1, 0, 0, 0, 0, time.UTC), "0 months / 0 years"},
		{"only excluded years", []salesRecord{prior}, time.Date(2026, time.March, 1, 0, 0, 0, 0, time.UTC), "0 months / 0 years"},
		{"report before cutoff", []salesRecord{prior}, time.Date(2025, time.December, 31, 0, 0, 0, 0, time.UTC), "0 months / 0 years"},
		{"one eligible completed month", []salesRecord{prior, {Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{10, 999}}}, time.Date(2026, time.February, 28, 0, 0, 0, 0, time.UTC), "1 months / 1 years"},
		{"eligible zeros before first positive shipment", []salesRecord{prior, {Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{0, 10}}}, time.Date(2026, time.March, 1, 0, 0, 0, 0, time.UTC), "1 months / 1 years"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			f := newProductLineWorkbook()
			t.Cleanup(func() { _ = f.Close() })
			item := &inventoryEntry{SKU: "BD1001FJ", ProductLine: "2021", OnHand: 30, SalesRecords: tc.records}
			if err := writeMTOSheet(f, []*inventoryEntry{item}, tc.asOf); err != nil {
				t.Fatalf("write MTO as of %s with history %+v: %v", tc.asOf, tc.records, err)
			}
			for _, check := range []struct{ cell, want string }{
				{"C2", ""}, {"D2", "Insufficient history"}, {"E2", "Insufficient history"}, {"I2", tc.coverage},
			} {
				if got, err := f.GetCellValue("MTO", check.cell); err != nil || got != check.want {
					t.Errorf("%s: expected %q, got %q (error %v)", check.cell, check.want, got, err)
				}
			}
		})
	}
}

// TestMTOCutoffPreservesMonthlyHistory verifies generating MTO does not remove
// or overwrite the pre-2026 records rendered by Monthly History afterward.
func TestMTOCutoffPreservesMonthlyHistory(t *testing.T) {
	item := &inventoryEntry{SKU: "BD1001FJ", ProductLine: "2021", OnHand: 30, SalesRecords: []salesRecord{
		{Year: 2025, Metric: "Quantity Sold", Periods: [12]float64{90}},
		{Year: 2025, Metric: "Quantity Issued", Periods: [12]float64{10}},
		{Year: 2026, Metric: "Quantity Sold", Periods: [12]float64{10, 20}},
	}}
	f := newProductLineWorkbook()
	t.Cleanup(func() { _ = f.Close() })
	if err := writeMTOSheet(f, []*inventoryEntry{item}, time.Date(2026, time.March, 1, 0, 0, 0, 0, time.UTC)); err != nil {
		t.Fatalf("write MTO for listed 2021 SKU: %v", err)
	}
	if err := writeMonthlyHistorySheet(f, []*inventoryEntry{item}); err != nil {
		t.Fatalf("write Monthly History after MTO: %v", err)
	}
	for _, check := range []struct{ cell, want string }{
		{"C2", "2025"}, {"D2", "Quantity Sold"}, {"E2", "90"},
		{"C3", "2025"}, {"D3", "Quantity Issued"}, {"E3", "10"},
		{"C4", "2026"}, {"D4", "Quantity Sold"}, {"E4", "10"}, {"F4", "20"},
	} {
		if got, err := f.GetCellValue("Monthly History", check.cell); err != nil || got != check.want {
			t.Errorf("Monthly History %s after MTO cutoff: expected %q, got %q (error %v)", check.cell, check.want, got, err)
		}
	}
}
