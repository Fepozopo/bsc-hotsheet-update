package hotsheet

import (
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

// TestLoadInventoryEntriesFooterBoundaries verifies that footer labels, dates,
// and times stop inventory import, while ordinary and empty SKU cells do not.
func TestLoadInventoryEntriesFooterBoundaries(t *testing.T) {
	for _, tc := range []struct {
		name    string
		cell    string
		wantSKU []string
	}{
		{name: "run date label", cell: "Run Date:", wantSKU: []string{"FIRST"}},
		{name: "case insensitive label", cell: "  run date  ", wantSKU: []string{"FIRST"}},
		{name: "short date", cell: "4/9/26", wantSKU: []string{"FIRST"}},
		{name: "long date", cell: "04/09/2026", wantSKU: []string{"FIRST"}},
		{name: "time", cell: "Printed 12:34", wantSKU: []string{"FIRST"}},
		{name: "ordinary SKU", cell: "BD1001FJ", wantSKU: []string{"FIRST", "BD1001FJ", "AFTER"}},
		{name: "empty SKU", wantSKU: []string{"FIRST", "AFTER"}},
		{name: "whitespace SKU", cell: "  ", wantSKU: []string{"FIRST", "AFTER"}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			t.Parallel()
			f := excelize.NewFile()
			t.Cleanup(func() { _ = f.Close() })
			// Sage places each SKU two rows before its inventory values.
			for _, cell := range [][2]string{
				{"B2", "FIRST"}, {"B4", "BAS"},
				{"B5", tc.cell}, {"B7", "BAS"},
				{"B8", "AFTER"}, {"B10", "BAS"},
			} {
				if err := f.SetCellStr("Sheet1", cell[0], cell[1]); err != nil {
					t.Fatalf("footer %q: cannot write %s=%q: %v", tc.cell, cell[0], cell[1], err)
				}
			}
			path := filepath.Join(t.TempDir(), "inventory.xlsx")
			if err := f.SaveAs(path); err != nil {
				t.Fatalf("footer %q: cannot save inventory: %v", tc.cell, err)
			}
			got, err := loadInventoryEntries(path, nil)
			if err != nil {
				t.Fatalf("footer %q: expected successful import of %v, got error %v", tc.cell, tc.wantSKU, err)
			}
			if len(got) != len(tc.wantSKU) {
				t.Errorf("footer %q: expected SKUs %v, got %v", tc.cell, tc.wantSKU, got)
			}
			for _, sku := range tc.wantSKU {
				if got[sku] == nil {
					t.Errorf("footer %q: expected SKU %q, got %v", tc.cell, sku, got)
				}
			}
		})
	}
}
