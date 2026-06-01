package hotsheet

import (
	"fmt"
	"log/slog"
	"strings"

	"github.com/xuri/excelize/v2"
)

var (
	poDataIdx          = colToIndex("A")
	poStatusIdx        = colToIndex("G")
	poOnPOIdx          = colToIndex("I")
	poOnPOBackorderIdx = colToIndex("K")
)

// mergePOData opens the optional PO workbook and merges PO numbers and quantities into the
// provided inventory map. PO-only SKUs are skipped so the workbook does not create UNKNOWN groups.
func mergePOData(poPath string, inventoryBySKU map[string]*inventoryEntry, logger *slog.Logger) error {
	if strings.TrimSpace(poPath) == "" {
		return nil
	}

	wbPO, err := excelize.OpenFile(poPath)
	if err != nil {
		return fmt.Errorf("failed to open PO report %s: %w", poPath, err)
	}
	defer func() {
		_ = wbPO.Close()
	}()

	poRows, err := wbPO.GetRows("Sheet1")
	if err != nil {
		return fmt.Errorf("failed to read PO sheet: %w", err)
	}
	if len(poRows) == 0 {
		return nil
	}

	for rowNum := 1; rowNum <= len(poRows); rowNum++ {
		row := poRows[rowNum-1]
		if poDataIdx >= len(row) {
			continue
		}

		sku := strings.TrimSpace(row[poDataIdx])
		if sku == "" {
			continue
		}

		item, ok := inventoryBySKU[sku]
		if !ok {
			if logger != nil {
				logger.Info("Skipping PO-only SKU (not present in inventory)", "SKU", sku)
			}
			continue
		}

		// Walk subsequent rows until we hit a line that starts with "Item" (end of section)
		// or we've collected up to two PO lines for this SKU.
		maxPOs := 2
		poCount := 0
		for scanRow := rowNum + 1; scanRow <= len(poRows) && poCount < maxPOs; scanRow++ {
			nextRow := getRow(poRows, scanRow)
			if nextRow == nil {
				break
			}

			poCell := strings.TrimSpace(getCell(nextRow, poDataIdx))
			if poCell == "" {
				// skip empty lines
				continue
			}
			if strings.HasPrefix(strings.ToUpper(poCell), "ITEM") {
				// end of PO block for this SKU
				break
			}

			poNum, qty := applyPOToEntry(item, nextRow)
			if logger != nil {
				status := strings.TrimSpace(getCell(nextRow, poStatusIdx))
				logger.Debug("Individual PO parse",
					"SKU", sku,
					"skuRow", rowNum,
					"poRow", scanRow,
					"PO Num", poNum,
					"QTY", qty,
					"PO Status", status,
				)
			}

			poCount++
		}
	}

	return nil
}

// applyPOToEntry normalizes one PO line, chooses the correct quantity column based on the status,
// and assigns the PO into the first available slot on the inventory entry.
func applyPOToEntry(item *inventoryEntry, poRow []string) (string, int) {
	status := strings.TrimSpace(getCell(poRow, poStatusIdx))
	var qty int
	if strings.EqualFold(status, "Back Order") {
		qty = parseInt(getCell(poRow, poOnPOBackorderIdx))
	} else {
		qty = parseInt(getCell(poRow, poOnPOIdx))
	}

	// Normalize PO numbers by removing leading zeros while preserving a usable zero value.
	poNum := strings.TrimLeft(strings.TrimSpace(getCell(poRow, poDataIdx)), "0")
	if poNum == "" {
		poNum = "0"
	}
	assignPO(item, poNum, qty)
	return poNum, qty
}

// assignPO assigns a PO number and quantity into the first available PO slot on the entry.
// The workbook only exposes two visible PO lines, so any additional PO quantities are folded
// into the first slot to avoid losing data.
func assignPO(item *inventoryEntry, poNum string, qty int) {
	if poNum == "" && qty == 0 {
		return
	}
	if item.PONum1 == "" {
		item.PONum1 = poNum
		item.OnPO1 = qty
		return
	}
	if item.PONum2 == "" {
		item.PONum2 = poNum
		item.OnPO2 = qty
		return
	}
	// fallback: accumulate into OnPO1
	item.OnPO1 += qty
}
