package hotsheet

import (
	"testing"

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
