package hotsheet

import (
	"reflect"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// TestStandardSheetsCommitted verifies the renamed column and its calculated
// quantity and availability with and without the optional PO detail columns.
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
			if err := writeStandardSheetHeaders(f, "Everyday", headers, tc.hasPO); err != nil {
				t.Fatalf("cannot write headers for %s: %v", tc.name, err)
			}
			item := &inventoryEntry{SKU: "A", Occasion: "EVERYDAY", OnHand: 12, OnPO: 9, OnSO: 3, OnBO: 4}
			if err := writeStandardSheetRows(f, "Everyday", []*inventoryEntry{item}, tc.hasPO, 6, mtoYtdIdx, mtoPyIdx); err != nil {
				t.Fatalf("cannot write rows for %s: %v", tc.name, err)
			}
			for cell, expected := range map[string]string{
				tc.committedCol + "1": "Quantity Committed",
				tc.committedCol + "2": "7",
				tc.availableCol + "2": "14",
			} {
				// RawCellValue checks the stored numeric output rather than Excel's display formatting.
				got, err := f.GetCellValue("Everyday", cell, excelize.Options{RawCellValue: true})
				if err != nil || got != expected {
					t.Errorf("%s cell %s: expected %q, got %q (error %v)", tc.name, cell, expected, got, err)
				}
			}
		})
	}
}

// TestMTOYTDSheet checks that each product-line workbook has a combined sheet after
// the seasonal tabs, with eligible products sorted by MTO YTD, a mapped Season column
// before Occasion, and otherwise unchanged seasonal rows.
func TestMTOYTDSheet(t *testing.T) {
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

			wantSheets := []string{"Everyday", "Winter", "Spring", "MTO‑YTD", "Data Insights", "Best Sellers"}
			if got := f.GetSheetList(); !reflect.DeepEqual(got, wantSheets) {
				t.Fatalf("hasPO=%v: expected sheets %v, got %v", tc.hasPO, wantSheets, got)
			}
			// GetRows reads the saved cell values, including the calculated MTO columns.
			combined, err := f.GetRows(mtoYTDSheetName, excelize.Options{RawCellValue: true})
			if err != nil {
				t.Fatalf("hasPO=%v: cannot read combined rows: %v", tc.hasPO, err)
			}
			// F has more stock than B but much faster YTD sales, so its MTO is lower.
			wantSKUs := []string{"D", "F", "B", "C", "E", "A"}
			seasonBySKU := map[string]string{"D": "Spring", "F": "Everyday", "B": "Winter", "C": "Spring", "E": "Winter", "A": "Everyday"}
			if len(combined) != len(wantSKUs)+1 {
				t.Fatalf("hasPO=%v: expected %d combined rows, got %d (%v)", tc.hasPO, len(wantSKUs)+1, len(combined), combined)
			}
			seasonal := make(map[string][][]string)
			for _, name := range []string{"Everyday", "Winter", "Spring"} {
				seasonal[name], err = f.GetRows(name, excelize.Options{RawCellValue: true})
				if err != nil {
					t.Fatalf("hasPO=%v: cannot read %s rows: %v", tc.hasPO, name, err)
				}
				if len(combined[0]) != len(seasonal[name][0])+1 || len(combined[0]) <= tc.seasonIndex+1 {
					t.Fatalf("hasPO=%v: expected %s header plus Season, got %v vs %v", tc.hasPO, name, combined[0], seasonal[name][0])
				}
				if combined[0][tc.seasonIndex] != "Season" || combined[0][tc.seasonIndex+1] != "Occasion" {
					t.Errorf("hasPO=%v: expected Season then Occasion at index %d, got %v", tc.hasPO, tc.seasonIndex, combined[0])
				}
				withoutSeason := append([]string(nil), combined[0][:tc.seasonIndex]...)
				withoutSeason = append(withoutSeason, combined[0][tc.seasonIndex+1:]...)
				if !reflect.DeepEqual(withoutSeason, seasonal[name][0]) {
					t.Errorf("hasPO=%v: expected %s header %v plus Season, got %v", tc.hasPO, name, seasonal[name][0], combined[0])
				}
			}
			for i, sku := range wantSKUs {
				row := combined[i+1]
				if len(row) == 0 || row[0] != sku {
					t.Fatalf("hasPO=%v: expected combined SKU %s at row %d, got %v", tc.hasPO, sku, i+2, row)
				}
				season := seasonBySKU[sku]
				var matched []string
				for _, seasonalRow := range seasonal[season][1:] {
					if len(seasonalRow) > 0 && seasonalRow[0] == sku {
						matched = seasonalRow
						break
					}
				}
				if len(row) <= tc.seasonIndex {
					t.Fatalf("hasPO=%v, SKU=%s: expected Season at index %d, got row %v", tc.hasPO, sku, tc.seasonIndex, row)
				}
				if row[tc.seasonIndex] != season {
					t.Errorf("hasPO=%v, SKU=%s: expected Season %s, got %q", tc.hasPO, sku, season, row[tc.seasonIndex])
				}
				withoutSeason := append([]string(nil), row[:tc.seasonIndex]...)
				withoutSeason = append(withoutSeason, row[tc.seasonIndex+1:]...)
				if !reflect.DeepEqual(withoutSeason, matched) {
					t.Errorf("hasPO=%v, SKU=%s: expected %s row %v plus Season, got combined row %v", tc.hasPO, sku, season, matched, row)
				}
			}
		})
	}
}
