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

// poLine holds a normalized PO number and its status-selected quantity.
type poLine struct {
	number   string
	quantity int
}

// poImportData groups parsed PO lines by matching inventory entry. Lines retain
// source order within each SKU; parsing never changes the inventory entries.
type poImportData map[*inventoryEntry][]poLine

// merge assigns the parsed lines to inventory using the existing two-slot policy.
// The caller must own the entries exclusively until merging finishes.
func (data poImportData) merge() {
	for item, lines := range data {
		for _, line := range lines {
			assignPO(item, line.number, line.quantity)
		}
	}
}

// readPOData parses at most two lines per SKU block into private buffers, without
// modifying inventory. inventoryBySKU must remain read-only during parsing.
// Empty paths return no data; unreadable workbooks return an error and no data.
func readPOData(poPath string, inventoryBySKU map[string]*inventoryEntry, logger *slog.Logger) (poImportData, error) {
	if strings.TrimSpace(poPath) == "" {
		return nil, nil
	}

	wbPO, err := excelize.OpenFile(poPath)
	if err != nil {
		return nil, fmt.Errorf("failed to open PO report %s: %w", poPath, err)
	}
	defer func() {
		_ = wbPO.Close()
	}()

	// Excelize materializes the rows so each SKU block can scan ahead to its POs.
	poRows, err := wbPO.GetRows("Sheet1")
	if err != nil {
		return nil, fmt.Errorf("failed to read PO sheet: %w", err)
	}
	data := make(poImportData)

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
		if !ok || item == nil {
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

			poNum, qty := parsePOLine(nextRow)
			data[item] = append(data[item], poLine{number: poNum, quantity: qty})
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

	return data, nil
}

// parsePOLine returns a normalized PO number and the quantity selected by status.
// Back Order uses its dedicated column; leading zeros collapse to a usable "0".
func parsePOLine(poRow []string) (string, int) {
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
