package hotsheet

import (
	"fmt"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

// TestGenerateProductLineOutputs verifies the public pipeline saves every line,
// returns paths in stable product-line order, and reports monotonic completion.
// A header-only inventory must return an empty, non-nil output list successfully.
func TestGenerateProductLineOutputs(t *testing.T) {
	// lineItem keeps the expected SKU and product line independent of filenames
	// or column metadata produced by the generator.
	type lineItem struct{ line, sku string }
	for _, tc := range []struct {
		name  string
		items []lineItem
		want  []lineItem
	}{
		{name: "no product lines"},
		{
			name:  "multiple lines",
			items: []lineItem{{"BAS", "B-SKU"}, {"OAT", "O-SKU"}, {"ASC", "A-SKU"}},
			want:  []lineItem{{"ASC", "A-SKU"}, {"BAS", "B-SKU"}, {"OAT", "O-SKU"}},
		},
	} {
		t.Run(tc.name, func(t *testing.T) {
			inventory := excelize.NewFile()
			t.Cleanup(func() { _ = inventory.Close() })
			if err := inventory.SetCellStr("Sheet1", "A1", "Inventory"); err != nil {
				t.Fatalf("write inventory header: %v", err)
			}
			for index, item := range tc.items {
				// Sage stores a SKU every three rows, two rows before its values.
				row := index*3 + 2
				for _, cell := range []struct{ name, value string }{
					{fmt.Sprintf("B%d", row), item.sku},
					{fmt.Sprintf("B%d", row+2), item.line},
					{fmt.Sprintf("D%d", row+2), "Counter Cards"},
					{fmt.Sprintf("F%d", row+2), "Carryover"},
					{fmt.Sprintf("H%d", row+2), "5"},
				} {
					if err := inventory.SetCellStr("Sheet1", cell.name, cell.value); err != nil {
						t.Fatalf("SKU %s: set %s=%q: %v", item.sku, cell.name, cell.value, err)
					}
				}
			}
			inputPath := filepath.Join(t.TempDir(), "inventory.xlsx")
			if err := inventory.SaveAs(inputPath); err != nil {
				t.Fatalf("save inventory fixture: %v", err)
			}
			var progress []Progress
			outputs, err := Generate(inputPath, "", "", "", t.TempDir(), nil, func(update Progress) {
				progress = append(progress, update)
			})
			if err != nil || outputs == nil || len(outputs) != len(tc.want) {
				t.Fatalf("items %v: expected %d non-nil output paths and no error, got %v and %v", tc.items, len(tc.want), outputs, err)
			}
			for index, path := range outputs {
				want := tc.want[index]
				if !strings.HasPrefix(filepath.Base(path), want.line+"_hotsheet_") {
					t.Errorf("output %d: expected product line %s, got %s", index, want.line, path)
				}
				written, err := excelize.OpenFile(path)
				if err != nil {
					t.Fatalf("open saved product line %s: %v", want.line, err)
				}
				t.Cleanup(func() { _ = written.Close() })
				if got, err := written.GetCellValue("All Products", "A2"); err != nil || got != want.sku {
					t.Errorf("line %s: expected SKU %s in A2, got %q (error %v)", want.line, want.sku, got, err)
				}
			}
			if len(progress) == 0 || progress[0].Percent != 0 || progress[len(progress)-1].Percent != 100 {
				t.Fatalf("expected progress from 0 to 100, got %v", progress)
			}
			previous := 0
			for _, update := range progress {
				if update.Percent < previous || update.Percent > 100 {
					t.Errorf("after %d%%: expected nondecreasing progress up to 100, got %+v", previous, update)
				}
				previous = update.Percent
			}
		})
	}
}
