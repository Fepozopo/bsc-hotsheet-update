package hotsheet

import (
	"errors"
	"math"
	"path/filepath"
	"strconv"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// writeIssueHistoryFixture creates a Sage-shaped report with BSC item blocks,
// other-warehouse data and subtotals. It returns the saved workbook path.
func writeIssueHistoryFixture(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()
	rows := map[int][]interface{}{
		1:  {"Item Code", "Period 1", "Period 2", "Period 3", "Period 4", "Period 5", "Period 6", "Period 7", "Period 8", "Period 9", "Period 10", "Period 11", "Period 12"},
		2:  {"SKU-1", "description", "Product Line:", "BAS"},
		3:  {"Warehouse:", "BSC  Biely Shoaf Seattle"},
		4:  {"Year:", 2025},
		5:  {"Quantity Issued:", 3, -7, 0, 0, 0, 0, 0, 0, 0, 0, 0, 9},
		6:  {"Quantity Transferred:", 100, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0},
		7:  {"Total For Warehouse:  BSC  Biely Shoaf Seattle"},
		8:  {"Quantity Issued:", 1000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0},
		9:  {"Warehouse:", "P3A  Page 3"},
		10: {"Year:", 2025},
		11: {"Quantity Issued:", 500, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0},
		12: {"Total For Item:  SKU-1"},
		13: {"Quantity Issued:", 1503, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0},
		14: {"OTHER", "description", "Product Line:", "BAS"},
		15: {"Warehouse:", "BSC  Biely Shoaf Seattle"},
		16: {"Year:", 2025},
		17: {"Quantity Issued:", 400, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0},
		18: {"Report Total:"},
		19: {"Quantity Issued:", 1903, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0},
		20: {"Run Date: 9/29/2026   6:31:47PM"},
	}
	for row, values := range rows {
		for col, value := range values {
			cell, err := excelize.CoordinatesToCellName(col+1, row)
			if err != nil {
				t.Fatalf("issue fixture row %d col %d: %v", row, col, err)
			}
			if err := f.SetCellValue("Sheet1", cell, value); err != nil {
				t.Fatalf("issue fixture cell %s: %v", cell, err)
			}
		}
	}
	path := filepath.Join(t.TempDir(), "issue.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save issue history fixture %s: %v", path, err)
	}
	return path
}

// TestMergeIssueHistory verifies BSC item-only import, month-level clamping,
// retention of sales and inventory data, and month-range shipment totals.
func TestMergeIssueHistory(t *testing.T) {
	item := &inventoryEntry{SKU: "SKU-1", YTDSold: 42, YTDIssued: 4, SalesRecords: []salesRecord{
		{Year: 2025, Metric: "Quantity Sold", Periods: [12]float64{2, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 5}},
		{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{20, 40, 0, 0, 0, 0, 0, 0, 0, 0, 0, 50}},
	}}
	if err := mergeIssueHistory(writeIssueHistoryFixture(t), map[string]*inventoryEntry{item.SKU: item}, nil); err != nil {
		t.Fatalf("issue history import for SKU-1: %v", err)
	}
	if len(item.SalesRecords) != 3 || item.SalesRecords[2].Metric != "Quantity Issued" || item.SalesRecords[2].Periods[0] != 3 || item.SalesRecords[2].Periods[1] != 0 || item.SalesRecords[2].Periods[11] != 9 || item.YTDSold != 42 || item.YTDIssued != 4 {
		t.Fatalf("expected one clamped BSC issue row without changing YTD sales for SKU-1, got %+v", item)
	}
	for _, tc := range []struct {
		name    string
		period  *BestSellersRange
		units   float64
		dollars float64
	}{
		{"inventory YTD", nil, 46, 0},
		{"January and February", &BestSellersRange{2025, 1, 2025, 2}, 9, 60},
		{"December only", &BestSellersRange{2025, 12, 2025, 12}, 14, 50},
		{"outside year", &BestSellersRange{2024, 1, 2024, 12}, 0, 0},
	} {
		t.Run(tc.name, func(t *testing.T) {
			got := bestSellerRows([]*inventoryEntry{item}, tc.period)[0]
			if got.shipped != tc.units || got.dollars != tc.dollars {
				t.Errorf("period %+v for SKU-1: expected shipped=%v dollars=%v, got shipped=%v dollars=%v", tc.period, tc.units, tc.dollars, got.shipped, got.dollars)
			}
		})
	}
}

// TestIssueHistoryRequiresSales validates the public generation boundary before
// trying to read inventory or produce any output files.
func TestIssueHistoryRequiresSales(t *testing.T) {
	paths, err := Generate("", "", "", "issue.xlsx", t.TempDir(), nil, nil)
	if !errors.Is(err, ErrIssueHistorySalesRequired) || paths != nil {
		t.Fatalf("issue without sales: expected ErrIssueHistorySalesRequired and no paths, got paths=%v err=%v", paths, err)
	}
}

// TestMTOUsesShippedMonths checks that sold and issued months merge before
// coverage/seasonal weighting; negatives, partial months and other metrics do not.
func TestMTOUsesShippedMonths(t *testing.T) {
	asOf := time.Date(2025, time.April, 15, 0, 0, 0, 0, time.UTC)
	records := []salesRecord{
		{Year: 2025, Metric: "Quantity Sold", Periods: [12]float64{2, 4, 0, 0}},
		{Year: 2025, Metric: "Quantity Issued", Periods: [12]float64{3, -7, 6, 999}},
		{Year: 2025, Metric: "Quantity Transferred", Periods: [12]float64{900}},
	}
	profile := buildMTOHistoryProfile(records, asOf)
	if !profile.usable || profile.months != 3 || profile.years != 1 || profile.monthly[0] != 5 || profile.monthly[1] != 4 || profile.monthly[2] != 6 || math.Abs(profile.monthly[3]-5) > 1e-9 {
		t.Fatalf("sold+issued completed Jan-Mar 2025: expected 3 months, Jan=5 Feb=4 Mar=6 Apr=5, got %+v", profile)
	}
}

// TestMTOShippedDemandSheet verifies the saved forecast consumes combined monthly
// shipments, not two independent observations for sold and issued quantities.
func TestMTOShippedDemandSheet(t *testing.T) {
	f := newProductLineWorkbook()
	defer func() { _ = f.Close() }()
	item := &inventoryEntry{SKU: "SKU-1", Status: "Active", OnHand: 100, SalesRecords: []salesRecord{
		{Year: 2025, Metric: "Quantity Sold", Periods: [12]float64{2, 4}},
		{Year: 2025, Metric: "Quantity Issued", Periods: [12]float64{3, -7, 6}},
	}}
	asOf := time.Date(2025, time.April, 1, 0, 0, 0, 0, time.UTC)
	if err := writeMTOSheet(f, []*inventoryEntry{item}, asOf); err != nil {
		t.Fatalf("write MTO with sold and issued history for SKU-1: %v", err)
	}
	units, err := f.GetCellValue(mtoSheetName, "C2", excelize.Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("MTO C2 for SKU-1: %v", err)
	}
	demand, err := strconv.ParseFloat(units, 64)
	if err != nil || math.Abs(demand-60) > 1e-9 {
		t.Errorf("MTO C2 for SKU-1: expected 60 shipped units, got %q (parse error %v)", units, err)
	}
	coverage, err := f.GetCellValue(mtoSheetName, "I2")
	if err != nil || coverage != "3 months / 1 years" {
		t.Errorf("MTO I2 for SKU-1: expected 3 months / 1 years, got %q (err %v)", coverage, err)
	}
}

// TestWriteShippedBestSellersHeader checks the shipped label for inventory YTD
// and ranged sales, including the issue-free case, with the appropriate source.
func TestWriteShippedBestSellersHeader(t *testing.T) {
	for _, tc := range []struct {
		name         string
		period       *BestSellersRange
		includeIssue bool
		want         string
	}{
		{"ranged with issues", &BestSellersRange{2025, 1, 2025, 1}, true, "5"},
		{"ranged without issues", &BestSellersRange{2025, 1, 2025, 1}, false, "2"},
		{"inventory YTD", nil, false, "12"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			f := newProductLineWorkbook()
			defer func() { _ = f.Close() }()
			item := &inventoryEntry{SKU: "SKU-1", YTDSold: 8, YTDIssued: 4, SalesRecords: []salesRecord{{Year: 2025, Metric: "Quantity Sold", Periods: [12]float64{2}}}}
			if tc.includeIssue {
				item.SalesRecords = append(item.SalesRecords, salesRecord{Year: 2025, Metric: "Quantity Issued", Periods: [12]float64{3}})
			}
			if err := writeBestSellersSheet(f, []*inventoryEntry{item}, tc.period); err != nil {
				t.Fatalf("write Best Sellers for %s: %v", tc.name, err)
			}
			header, err := f.GetCellValue(bestSellersSheetName, "C1")
			if err != nil || header != "Quantity Shipped" {
				t.Errorf("%s C1: expected %q, got %q (err %v)", tc.name, "Quantity Shipped", header, err)
			}
			units, err := f.GetCellValue(bestSellersSheetName, "C2")
			if err != nil || units != tc.want {
				t.Errorf("%s C2: expected %s, got %s (err %v)", tc.name, tc.want, units, err)
			}
		})
	}
}
