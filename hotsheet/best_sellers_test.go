package hotsheet

import (
	"archive/zip"
	"errors"
	"io"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

// TestBestSellersRangeValidate verifies boundary months and reversed or invalid years.
func TestBestSellersRangeValidate(t *testing.T) {
	cases := []struct {
		name    string
		rangeIn BestSellersRange
		wantErr bool
	}{
		{"same month", BestSellersRange{2024, 1, 2024, 1}, false},
		{"cross year", BestSellersRange{2024, 12, 2025, 1}, false},
		{"zero month", BestSellersRange{2024, 0, 2024, 12}, true},
		{"thirteenth month", BestSellersRange{2024, 1, 2024, 13}, true},
		{"year zero", BestSellersRange{0, 1, 2024, 1}, true},
		{"year beyond Excel range", BestSellersRange{2024, 1, 10000, 1}, true},
		{"reversed months", BestSellersRange{2024, 5, 2024, 4}, true},
		{"reversed years", BestSellersRange{2025, 1, 2024, 12}, true},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			err := tc.rangeIn.Validate()
			if (err != nil) != tc.wantErr {
				t.Errorf("range %+v: expected error=%v, got %v", tc.rangeIn, tc.wantErr, err)
			}
		})
	}
}

// TestBestSellerRows checks inventory YTD shipments, inclusive monthly totals,
// zero-shipment items, and deterministic order without changing input order.
func TestBestSellerRows(t *testing.T) {
	items := []*inventoryEntry{
		{SKU: "B", YTDSold: 10, YTDIssued: 2, DollarSoldYTD: 100, SalesRecords: []salesRecord{
			{Year: 2024, Metric: "Quantity Sold", Periods: [12]float64{0, 8}},
			{Year: 2024, Metric: "Dollars Sold", Periods: [12]float64{0, 80}},
			{Year: 2025, Metric: "Quantity Sold", Periods: [12]float64{1}},
			{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{10}},
		}},
		{SKU: "A", YTDSold: 3, YTDIssued: 3, DollarSoldYTD: 30, SalesRecords: []salesRecord{
			{Year: 2024, Metric: "Quantity Sold", Periods: [12]float64{1, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3}},
			{Year: 2024, Metric: "Dollars Sold", Periods: [12]float64{10, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 30}},
			{Year: 2025, Metric: "Quantity Sold", Periods: [12]float64{4, 5, 6}},
			{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{40, 50, 60}},
			{Year: 2025, Metric: "Quantity Returned", Periods: [12]float64{999}},
		}},
		{SKU: "C", YTDSold: 0, YTDIssued: -5, DollarSoldYTD: 0},
	}
	cases := []struct {
		name   string
		period *BestSellersRange
		order  [3]string
		quant  [3]float64
		dollar [3]float64
	}{
		{"inventory YTD shipped despite history", nil, [3]string{"B", "A", "C"}, [3]float64{12, 6, 0}, [3]float64{100, 30, 0}},
		{"inclusive same-year endpoints", &BestSellersRange{2025, 1, 2025, 2}, [3]string{"A", "B", "C"}, [3]float64{9, 1, 0}, [3]float64{90, 10, 0}},
		{"inclusive cross-year endpoints", &BestSellersRange{2024, 2, 2025, 2}, [3]string{"A", "B", "C"}, [3]float64{14, 9, 0}, [3]float64{140, 90, 0}},
		{"no records in period", &BestSellersRange{2026, 1, 2026, 12}, [3]string{"A", "B", "C"}, [3]float64{0, 0, 0}, [3]float64{0, 0, 0}},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			rows := bestSellerRows(items, tc.period)
			if len(rows) != len(tc.order) {
				t.Fatalf("period %+v: expected %d rows, got %d", tc.period, len(tc.order), len(rows))
			}
			for i, row := range rows {
				if row.item.SKU != tc.order[i] || row.shipped != tc.quant[i] || row.dollars != tc.dollar[i] {
					t.Errorf("period %+v row %d: expected SKU=%s quantity=%v dollars=%v, got SKU=%s quantity=%v dollars=%v", tc.period, i, tc.order[i], tc.quant[i], tc.dollar[i], row.item.SKU, row.shipped, row.dollars)
				}
			}
			if items[0].SKU != "B" || items[1].SKU != "A" {
				t.Errorf("period %+v: input order changed to %s, %s; expected B, A", tc.period, items[0].SKU, items[1].SKU)
			}
		})
	}
}

// TestWriteBestSellersSheet verifies ranked YTD shipments, inventory snapshot quantities,
// metadata, status, and filter dropdowns in a saved workbook.
func TestWriteBestSellersSheet(t *testing.T) {
	f := newProductLineWorkbook()
	defer func() { _ = f.Close() }()
	items := []*inventoryEntry{
		{SKU: "B", Description: "second", YTDSold: 1, YTDIssued: -5, DollarSoldYTD: 20, Status: "Rundown"},
		{SKU: "A", Description: "first", RawClassDesc: "Counter Cards", ClassDesc: "BX - Counter Cards", RoyaltyCode: "HOUSE", Occasion: "BIRTHDAY", Foil: "Yes", YTDSold: 5, YTDIssued: 3, DollarSoldYTD: 55.5, OnHand: 12, OnSO: 3, OnBO: 4, OnPO: 9, Status: "Carryover"},
	}
	if err := writeBestSellersSheet(f, items, nil); err != nil {
		t.Fatalf("cannot write Best Sellers sheet: %v", err)
	}
	path := filepath.Join(t.TempDir(), "best-sellers.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("cannot save Best Sellers workbook: %v", err)
	}
	written, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("cannot reopen Best Sellers workbook: %v", err)
	}
	defer func() { _ = written.Close() }()
	want := map[string]string{
		"A1": "Item Code", "B1": "Description", "C1": "Quantity Shipped", "D1": "Dollars Sold",
		"E1": "Quantity on Hand", "F1": "Quantity Committed", "G1": "Quantity on PO",
		"H1": "Royalty Code", "I1": "Class Description", "J1": "Occasion", "K1": "Foil Status", "L1": "Status",
		"A2": "A", "B2": "first", "C2": "8", "D2": "55.5", "E2": "12", "F2": "7", "G2": "9", "H2": "HOUSE",
		"I2": "Counter Cards", "J2": "BIRTHDAY", "K2": "Yes", "L2": "Carryover",
		"A3": "B", "B3": "second", "C3": "1", "D3": "20", "E3": "0", "F3": "0", "G3": "0", "L3": "Rundown",
	}
	for cell, expected := range want {
		got, err := written.GetCellValue(bestSellersSheetName, cell, excelize.Options{RawCellValue: true})
		if err != nil || got != expected {
			t.Errorf("cell %s: expected %q, got %q (error %v)", cell, expected, got, err)
		}
	}

	archive, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("cannot inspect Best Sellers workbook: %v", err)
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
		if strings.Contains(string(contents), `<autoFilter ref="$A$1:$L$1"`) {
			found = true
		}
	}
	if !found {
		t.Fatal("expected an A1:L1 autofilter in the saved Best Sellers sheet")
	}
}

// TestGenerateRejectsRangeWithoutHistory ensures callers cannot silently get YTD
// values when they request monthly history without supplying its source report.
func TestGenerateRejectsRangeWithoutHistory(t *testing.T) {
	period := &BestSellersRange{FromYear: 2024, FromMonth: 1, ToYear: 2024, ToMonth: 12}
	paths, err := Generate("", "", "", "", t.TempDir(), period, nil)
	if !errors.Is(err, ErrBestSellersHistoryRequired) || paths != nil {
		t.Fatalf("range without history: expected history-required error and no files, got paths=%v error=%v", paths, err)
	}
}
