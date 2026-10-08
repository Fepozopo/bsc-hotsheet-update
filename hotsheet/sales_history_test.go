package hotsheet

import (
	"archive/zip"
	"errors"
	"io"
	"path/filepath"
	"slices"
	"strings"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// writeSalesHistoryFixture writes a minimal report-style workbook with two warehouses,
// two years, and item totals to exercise the same boundaries as the Sage export.
func writeSalesHistoryFixture(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()
	rows := map[int][]interface{}{
		1:  {"Item Code", "Period 1", "Period 2", "Period 3", "Period 4", "Period 5", "Period 6", "Period 7", "Period 8", "Period 9", "Period 10", "Period 11", "Period 12"},
		2:  {"SKU-1", "untrusted description", "Product Line:", "BAS"},
		3:  {"Warehouse: ", "BSC  Biely Shoaf Seattle"},
		4:  {"Year:", 2024},
		5:  {"Quantity Sold:", 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 12},
		6:  {"Dollars Sold:", 25.5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 50},
		7:  {"Gross Profit Percent:", 85.5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 90},
		8:  {"Cost of Goods Sold:", 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 5},
		9:  {"Quantity Returned:", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1},
		10: {"Year:", 2025},
		11: {"Quantity Sold:", 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 8},
		12: {"Warehouse: ", "P3A  Page 3"},
		13: {"Year:", 2024},
		14: {"Quantity Sold:", 999, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 999},
		15: {"Total For Item:  SKU-1  untrusted description"},
		16: {"Quantity Sold:", 1009, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1019},
		17: {"OTHER", "unmatched", "Product Line:", "BAS"},
		18: {"Warehouse: ", "BSC  Biely Shoaf Seattle"},
		19: {"Year:", 2024},
		20: {"Quantity Sold:", 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10},
		21: {"Report Total:"},
		22: {"Quantity Sold:", 5000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6000},
		23: {"Run Date: 9/25/2026   3:58:16PM"},
	}
	for rowNum := 1; rowNum <= len(rows); rowNum++ {
		for col, value := range rows[rowNum] {
			cell, err := excelize.CoordinatesToCellName(col+1, rowNum)
			if err != nil {
				t.Fatal(err)
			}
			if err := f.SetCellValue("Sheet1", cell, value); err != nil {
				t.Fatalf("cannot write fixture cell %s: %v", cell, err)
			}
		}
	}
	path := filepath.Join(t.TempDir(), "history.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("cannot save sales history fixture: %v", err)
	}
	return path
}

// TestMergeSalesHistory verifies that only matching BSC rows are attached to an
// inventory item and its authoritative inventory fields remain unchanged.
func TestMergeSalesHistory(t *testing.T) {
	item := &inventoryEntry{SKU: "SKU-1", Description: "inventory description", Status: "Rundown", YTDSold: 42}
	bySKU := map[string]*inventoryEntry{item.SKU: item}
	runDate, err := mergeSalesHistory(writeSalesHistoryFixture(t), bySKU, nil)
	if err != nil {
		t.Fatalf("unexpected sales history import failure: %v", err)
	}
	if expected := time.Date(2026, time.September, 25, 0, 0, 0, 0, time.UTC); !runDate.Equal(expected) {
		t.Fatalf("expected report run date %v, got %v", expected, runDate)
	}
	if item.Description != "inventory description" || item.YTDSold != 42 || item.Status != "Rundown" {
		t.Fatalf("inventory fields changed during import for SKU-1: %+v", item)
	}
	if len(bySKU) != 1 || len(item.SalesRecords) != 6 {
		t.Fatalf("expected one inventory SKU with six BSC year/metric rows, got %d SKUs and %d rows", len(bySKU), len(item.SalesRecords))
	}
	cases := []struct {
		index, year int
		metric      string
		jan, dec    float64
	}{
		{0, 2024, "Quantity Sold", 4, 12},
		{1, 2024, "Dollars Sold", 25.5, 50},
		{2, 2024, "Gross Profit Percent", 85.5, 90},
		{3, 2024, "Cost of Goods Sold", 3, 5},
		{4, 2024, "Quantity Returned", 0, 1},
		{5, 2025, "Quantity Sold", 6, 8},
	}
	for _, tc := range cases {
		record := item.SalesRecords[tc.index]
		if record.Year != tc.year || record.Metric != tc.metric || record.Periods[0] != tc.jan || record.Periods[11] != tc.dec {
			t.Errorf("record %d: expected %d %s Jan=%v Dec=%v, got %+v", tc.index, tc.year, tc.metric, tc.jan, tc.dec, record)
		}
	}
}

// TestMergeSalesHistoryOptionalAndInvalid checks the optional no-op and rejects a
// malformed month instead of importing a silently incomplete sales record.
func TestMergeSalesHistoryOptionalAndInvalid(t *testing.T) {
	item := &inventoryEntry{SKU: "SKU-1"}
	bySKU := map[string]*inventoryEntry{"SKU-1": item}
	if date, err := mergeSalesHistory("", bySKU, nil); err != nil || !date.IsZero() || len(item.SalesRecords) != 0 {
		t.Fatalf("empty optional history: expected no date, records, or error; got records=%v date=%v error=%v", item.SalesRecords, date, err)
	}
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()
	for cell, value := range map[string]interface{}{"A1": "Item Code", "B1": "Period 1", "M1": "Period 12", "A2": "SKU-1", "C2": "Product Line:", "A3": "Warehouse:", "B3": "BSC", "A4": "Year:", "B4": 2024, "A5": "Quantity Sold:", "B5": 2} {
		if err := f.SetCellValue("Sheet1", cell, value); err != nil {
			t.Fatalf("cannot set fixture cell %s: %v", cell, err)
		}
	}
	path := filepath.Join(t.TempDir(), "invalid.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("cannot save invalid history fixture: %v", err)
	}
	_, err := mergeSalesHistory(path, bySKU, nil)
	if !errors.Is(err, errMissingSalesPeriod) {
		t.Fatalf("expected a missing sales-period error for invalid month, got %v", err)
	}
	if len(item.SalesRecords) != 0 {
		t.Fatalf("invalid row must not append a record, got %+v", item.SalesRecords)
	}
}

// TestMergeSalesHistoryRequiresRunDate ensures a report without a dated footer
// cannot anchor an MTO forecast to an invented current date.
func TestMergeSalesHistoryRequiresRunDate(t *testing.T) {
	file := excelize.NewFile()
	defer func() { _ = file.Close() }()
	for cell, value := range map[string]string{"A1": "Item Code", "B1": "Period 1", "M1": "Period 12"} {
		if err := file.SetCellValue("Sheet1", cell, value); err != nil {
			t.Fatalf("cannot create run-date fixture at %s: %v", cell, err)
		}
	}
	path := filepath.Join(t.TempDir(), "no-run-date.xlsx")
	if err := file.SaveAs(path); err != nil {
		t.Fatalf("cannot save run-date fixture: %v", err)
	}
	date, err := mergeSalesHistory(path, map[string]*inventoryEntry{}, nil)
	if !errors.Is(err, errMissingSalesHistoryRunDate) || !date.IsZero() {
		t.Fatalf("no run date: expected run-date error and zero date, got date=%v error=%v", date, err)
	}
}

// TestMonthlyHistorySheetOptional verifies the saved tab order: Best Sellers follows
// Data Insights, with Monthly History and MTO appended only when history is supplied.
func TestMonthlyHistorySheetOptional(t *testing.T) {
	for _, tc := range []struct {
		name       string
		hasHistory bool
		wantSheet  bool
	}{
		{"omitted", false, false},
		{"supplied", true, true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			path, err := buildProductLineWorkbook("BAS", nil, t.TempDir(), "20260101", false, tc.hasHistory, nil, time.Date(2026, time.September, 25, 0, 0, 0, 0, time.UTC), nil)
			if err != nil {
				t.Fatalf("hasHistory=%v: cannot build workbook: %v", tc.hasHistory, err)
			}
			file, err := excelize.OpenFile(path)
			if err != nil {
				t.Fatalf("hasHistory=%v: cannot read workbook: %v", tc.hasHistory, err)
			}
			t.Cleanup(func() { _ = file.Close() })
			wantSheets := []string{"All Products", "YTD Stock Priority", "Data Insights", "Best Sellers"}
			if tc.wantSheet {
				wantSheets = append(wantSheets, "Monthly History", "MTO", "MTO Forecast Data")
			}
			if got := file.GetSheetList(); !slices.Equal(got, wantSheets) {
				t.Errorf("hasHistory=%v: expected sheet order %v, got %v", tc.hasHistory, wantSheets, got)
			}
		})
	}
}

// TestWriteMonthlyHistorySheet verifies the user-visible row layout and Excel
// autofilter in a saved workbook, including the last Status column.
func TestWriteMonthlyHistorySheet(t *testing.T) {
	f := newProductLineWorkbook()
	defer func() { _ = f.Close() }()
	item := &inventoryEntry{SKU: "SKU-1", Description: "inventory description", Status: "Discontinued", SalesRecords: []salesRecord{{Year: 2024, Metric: "Dollars Sold", Periods: [12]float64{25.5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 50}}}}
	if err := writeMonthlyHistorySheet(f, []*inventoryEntry{item}); err != nil {
		t.Fatalf("cannot write Monthly History sheet: %v", err)
	}
	path := filepath.Join(t.TempDir(), "output.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("cannot save Monthly History workbook: %v", err)
	}
	written, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("cannot reopen Monthly History workbook: %v", err)
	}
	defer func() { _ = written.Close() }()
	for cell, want := range map[string]string{"A1": "Item Code", "E1": "Jan", "P1": "Dec", "Q1": "Status", "A2": "SKU-1", "B2": "inventory description", "C2": "2024", "D2": "Dollars Sold", "E2": "25.5", "P2": "50", "Q2": "Discontinued"} {
		got, err := written.GetCellValue(monthlyHistorySheetName, cell, excelize.Options{RawCellValue: true})
		if err != nil || got != want {
			t.Errorf("cell %s: expected %q, got %q (error %v)", cell, want, got, err)
		}
	}
	archive, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("cannot inspect output workbook: %v", err)
	}
	defer func() { _ = archive.Close() }()
	// Excel stores the filter as sheet XML; verify the saved file contains it.
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
		if strings.Contains(string(contents), `<autoFilter ref="$A$1:$Q$1"`) {
			found = true
		}
	}
	if !found {
		t.Fatal("expected an A1:Q1 autofilter in the saved workbook, but none was found")
	}
}
