package hotsheet

import (
	"reflect"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// TestStandardSheetsCommitted verifies All Products' committed quantity and availability
// with and without the optional PO detail columns.
func TestStandardSheetsCommitted(t *testing.T) {
	cases := []struct {
		name         string
		hasPO        bool
		committedCol string
		availableCol string
	}{
		{name: "without PO details", committedCol: "D", availableCol: "E"},
		{name: "with PO details", hasPO: true, committedCol: "H", availableCol: "I"},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			f := newProductLineWorkbook()
			defer func() { _ = f.Close() }()
			headers, mtoYtdIdx, mtoPyIdx := buildStandardSheetHeaders(tc.hasPO)
			if err := writeStandardSheetHeaders(f, allProductsSheetName, headers, tc.hasPO); err != nil {
				t.Fatalf("cannot write headers for %s: %v", tc.name, err)
			}
			item := &inventoryEntry{SKU: "A", Occasion: "EVERYDAY", OnHand: 12, OnPO: 9, OnSO: 3, OnBO: 4}
			if err := writeStandardSheetRows(f, allProductsSheetName, []*inventoryEntry{item}, tc.hasPO, 6, mtoYtdIdx, mtoPyIdx); err != nil {
				t.Fatalf("cannot write rows for %s: %v", tc.name, err)
			}
			for cell, expected := range map[string]string{
				tc.committedCol + "1": "Quantity Committed",
				tc.committedCol + "2": "7",
				tc.availableCol + "2": "14",
			} {
				// RawCellValue checks the stored numeric output rather than Excel's display formatting.
				got, err := f.GetCellValue(allProductsSheetName, cell, excelize.Options{RawCellValue: true})
				if err != nil || got != expected {
					t.Errorf("%s cell %s: expected %q, got %q (error %v)", tc.name, cell, expected, got, err)
				}
			}
		})
	}
}

// TestStandardSheetsShipped verifies that both inventory tabs label combined
// sold-plus-issued YTD/PY quantities as shipped and clamp negative issued values.
func TestStandardSheetsShipped(t *testing.T) {
	for _, tc := range []struct {
		name      string
		issuedYTD int
		issuedPY  int
		wantYTD   string
		wantPY    string
	}{
		{"positive issues", 3, 2, "10", "6"},
		{"negative issues", -5, -3, "7", "4"},
		{"zero issues", 0, 0, "7", "4"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			f := newProductLineWorkbook()
			defer func() { _ = f.Close() }()
			item := &inventoryEntry{SKU: "SKU-1", YTDSold: 7, YTDIssued: tc.issuedYTD, SoldPY: 4, IssuedPY: tc.issuedPY}
			if err := writeStandardSheets(f, []*inventoryEntry{item}, false); err != nil {
				t.Fatalf("write standard sheets for %s: %v", tc.name, err)
			}
			for _, sheet := range standardSheetNames {
				for cell, want := range map[string]string{"H1": "QTY Shipped YTD", "I1": "QTY Shipped PY", "H2": tc.wantYTD, "I2": tc.wantPY} {
					got, err := f.GetCellValue(sheet, cell, excelize.Options{RawCellValue: true})
					if err != nil || got != want {
						t.Errorf("%s %s for %s: expected %q, got %q (err %v)", sheet, cell, tc.name, want, got, err)
					}
				}
			}
		})
	}
}

// TestInventorySheets checks the saved workbook's two inventory tabs: All Products
// retains every input row and YTD Stock Priority excludes retired items and sorts by MTO.
// Both tabs show the original season mapping before Occasion with and without PO details.
func TestInventorySheets(t *testing.T) {
	for _, tc := range []struct {
		name        string
		hasPO       bool
		seasonIndex int
	}{
		{name: "without PO details", seasonIndex: 11},
		{name: "with PO details", hasPO: true, seasonIndex: 15},
	} {
		t.Run(tc.name, func(t *testing.T) {
			entries := []*inventoryEntry{
				{SKU: "C", Occasion: "EASTER", Status: "Active", OnHand: 15, YTDSold: 10},
				{SKU: "A", Occasion: "BIRTHDAY", Status: "Carryover", OnHand: 30, YTDSold: 10},
				{SKU: "R", Occasion: "BIRTHDAY", Status: "Rundown", OnHand: -100},
				{SKU: "E", Occasion: "CHRISTMAS", OnHand: 15, YTDSold: 10},
				{SKU: "B", Occasion: "CHRISTMAS", Status: "Active", OnHand: 5, YTDSold: 10},
				{SKU: "F", Status: "Active", OnHand: 40, YTDSold: 1000},
				{SKU: "D", Occasion: "EASTER", Status: "Active"},
				{SKU: "X", Occasion: "CHRISTMAS", Status: "Discontinued", OnHand: -100},
			}
			path, err := buildProductLineWorkbook("BAS", entries, t.TempDir(), "20260101", tc.hasPO, false, nil, time.Time{}, nil)
			if err != nil {
				t.Fatalf("hasPO=%v: cannot build workbook: %v", tc.hasPO, err)
			}
			f, err := excelize.OpenFile(path)
			if err != nil {
				t.Fatalf("hasPO=%v: cannot open workbook %s: %v", tc.hasPO, path, err)
			}
			t.Cleanup(func() { _ = f.Close() })

			wantSheets := []string{"All Products", "YTD Stock Priority", "Data Insights", "Best Sellers"}
			if got := f.GetSheetList(); !reflect.DeepEqual(got, wantSheets) {
				t.Fatalf("hasPO=%v: expected sheets %v, got %v", tc.hasPO, wantSheets, got)
			}
			// GetRows reads the stored values, including the calculated MTO columns.
			allRows, err := f.GetRows(allProductsSheetName, excelize.Options{RawCellValue: true})
			if err != nil {
				t.Fatalf("hasPO=%v: cannot read All Products: %v", tc.hasPO, err)
			}
			priorityRows, err := f.GetRows(ytdStockPrioritySheetName, excelize.Options{RawCellValue: true})
			if err != nil {
				t.Fatalf("hasPO=%v: cannot read YTD Stock Priority: %v", tc.hasPO, err)
			}
			if len(allRows) != len(entries)+1 {
				t.Fatalf("hasPO=%v: expected %d All Products rows, got %d (%v)", tc.hasPO, len(entries)+1, len(allRows), allRows)
			}
			if len(priorityRows) == 0 {
				t.Fatalf("hasPO=%v: expected YTD Stock Priority header, got no rows", tc.hasPO)
			}
			if !reflect.DeepEqual(allRows[0], priorityRows[0]) || len(allRows[0]) <= tc.seasonIndex+1 ||
				allRows[0][tc.seasonIndex] != "Season" || allRows[0][tc.seasonIndex+1] != "Occasion" {
				t.Fatalf("hasPO=%v: expected identical headers with Season before Occasion at index %d, got %v and %v", tc.hasPO, tc.seasonIndex, allRows[0], priorityRows[0])
			}
			seasonBySKU := map[string]string{
				"C": "Spring", "A": "Everyday", "R": "Everyday", "E": "Winter",
				"B": "Winter", "F": "Everyday", "D": "Spring", "X": "Winter",
			}
			allBySKU := make(map[string][]string, len(entries))
			for i, item := range entries {
				row := allRows[i+1]
				if len(row) <= tc.seasonIndex || row[0] != item.SKU || row[tc.seasonIndex] != seasonBySKU[item.SKU] {
					t.Fatalf("hasPO=%v: expected SKU %s / Season %s at All Products row %d, got %v", tc.hasPO, item.SKU, seasonBySKU[item.SKU], i+2, row)
				}
				allBySKU[item.SKU] = row
			}
			// F has more stock than B but much faster YTD sales, so its MTO is lower.
			wantPrioritySKUs := []string{"D", "F", "B", "C", "E", "A"}
			if len(priorityRows) != len(wantPrioritySKUs)+1 {
				t.Fatalf("hasPO=%v: expected %d priority rows, got %d (%v)", tc.hasPO, len(wantPrioritySKUs)+1, len(priorityRows), priorityRows)
			}
			for i, sku := range wantPrioritySKUs {
				row := priorityRows[i+1]
				if len(row) == 0 || row[0] != sku || !reflect.DeepEqual(row, allBySKU[sku]) {
					t.Errorf("hasPO=%v: expected YTD Stock Priority row %d to match All Products SKU %s (%v), got %v", tc.hasPO, i+2, sku, allBySKU[sku], row)
				}
			}
		})
	}
}
