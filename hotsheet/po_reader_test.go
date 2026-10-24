package hotsheet

import (
	"fmt"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

// writePOFixture creates a Sage-shaped PO report covering status-selected
// quantities, leading zeros, repeated SKU blocks, and more than two PO lines.
// It returns a temporary workbook path managed by the test framework.
func writePOFixture(t *testing.T) string {
	t.Helper()
	file := excelize.NewFile()
	t.Cleanup(func() { _ = file.Close() })
	rows := [][]any{
		{"SKU-1"},
		{"000123", nil, nil, nil, nil, nil, "Open", nil, 5, nil, 999},
		{"000045", nil, nil, nil, nil, nil, "bAcK oRdEr", nil, 999, nil, 7},
		{"000999", nil, nil, nil, nil, nil, "Open", nil, 500},
		{"Item Total:"},
		{"OTHER"},
		{"000111", nil, nil, nil, nil, nil, "Open", nil, 11},
		{"SKU-1"},
		{"000000", nil, nil, nil, nil, nil, "Open", nil, 3},
		{"Item Total:"},
		{"ZERO"},
		{"000000", nil, nil, nil, nil, nil, "Open", nil, 0},
		{"Item Total:"},
		{"NIL"},
		{"000888", nil, nil, nil, nil, nil, "Open", nil, 8},
	}
	for index, row := range rows {
		// Excelize writes the full row, retaining blank columns so status and
		// the normal/back-order quantities occupy G, I, and K respectively.
		if err := file.SetSheetRow("Sheet1", fmt.Sprintf("A%d", index+1), &row); err != nil {
			t.Fatalf("PO fixture row %d: expected successful write, got %v", index+1, err)
		}
	}
	path := filepath.Join(t.TempDir(), "po.xlsx")
	if err := file.SaveAs(path); err != nil {
		t.Fatalf("save PO fixture %s: %v", path, err)
	}
	return path
}

// TestReadPOData verifies parsing cannot mutate inventory while other readers
// are active, and merging preserves the two-slot and per-block limit policies.
func TestReadPOData(t *testing.T) {
	item := &inventoryEntry{SKU: "SKU-1", OnPO: 99}
	zero := &inventoryEntry{SKU: "ZERO"}
	inventory := map[string]*inventoryEntry{"SKU-1": item, "ZERO": zero, "NIL": nil}
	data, err := readPOData(writePOFixture(t), inventory, nil)
	if err != nil {
		t.Fatalf("matching inventory: expected successful PO parse, got %v", err)
	}
	if len(inventory) != 3 || item.PONum1 != "" || item.PONum2 != "" || item.OnPO1 != 0 || item.OnPO2 != 0 || zero.PONum1 != "" {
		t.Fatalf("before merging: expected untouched inventory without PO-only entries, got %v, item %+v, zero %+v", inventory, item, zero)
	}
	data.merge()
	if item.PONum1 != "123" || item.OnPO1 != 8 || item.PONum2 != "45" || item.OnPO2 != 7 || item.OnPO != 99 {
		t.Errorf("SKU-1 after merge: expected PO1=123/8, PO2=45/7, inventory OnPO=99, got %+v", item)
	}
	if zero.PONum1 != "0" || zero.OnPO1 != 0 || zero.PONum2 != "" {
		t.Errorf("ZERO after merge: expected PO1=0/0 and no second PO, got %+v", zero)
	}
	if _, exists := inventory["OTHER"]; exists {
		t.Error("PO-only OTHER: expected no new inventory entry")
	}
}

// TestReadPODataEmptyInputs verifies optional empty paths, nil maps, empty maps,
// and nil inventory entries produce no assignments or errors.
func TestReadPODataEmptyInputs(t *testing.T) {
	path := writePOFixture(t)
	for _, tc := range []struct {
		name, path string
		inventory  map[string]*inventoryEntry
	}{
		{name: "omitted path"},
		{name: "whitespace path", path: " \t"},
		{name: "nil inventory", path: path},
		{name: "empty inventory", path: path, inventory: map[string]*inventoryEntry{}},
		{name: "nil item", path: path, inventory: map[string]*inventoryEntry{"SKU-1": nil}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			data, err := readPOData(tc.path, tc.inventory, nil)
			if err != nil || len(data) != 0 {
				t.Errorf("path %q inventory %v: expected no PO data and no error, got %v and %v", tc.path, tc.inventory, data, err)
			}
			data.merge()
		})
	}
}
