package hotsheet

import (
	"fmt"
	"log/slog"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Shared workbook column lookups keep the row-parsing helpers aligned with the source report layout.
var (
	inventorySKUIdx         = colToIndex("B")
	inventoryProductLineIdx = colToIndex("B")
	inventoryClassIdx       = colToIndex("D")
	inventoryStatusIdx      = colToIndex("F")
	inventoryOnHandIdx      = colToIndex("H")
	inventoryOnPOIdx        = colToIndex("J")
	inventoryOnSOIdx        = colToIndex("L")
	inventoryOnBOIdx        = colToIndex("N")
	inventoryTotalAvailIdx  = colToIndex("P")
	inventoryYTDSoldIdx     = colToIndex("R")
	inventoryYTDIssuedIdx   = colToIndex("T")
	inventorySoldPYIdx      = colToIndex("V")
	inventoryIssuedPYIdx    = colToIndex("X")
	inventoryFoilIdx        = colToIndex("Z")
	inventoryOccasionIdx    = colToIndex("AB")
	inventoryDescIdx        = colToIndex("AD")
	inventoryUPCIdx         = colToIndex("AF")
	inventoryRoyaltyCodeIdx = colToIndex("AH")
	inventoryDollarYTDIdx   = colToIndex("AJ")
	inventoryDollarPYIdx    = colToIndex("AL")
)

// loadInventoryEntries opens the inventory workbook, parses the inventory rows, and returns
// the populated inventory map keyed by SKU.
func loadInventoryEntries(inventoryPath string, logger *slog.Logger) (map[string]*inventoryEntry, error) {
	if logger != nil {
		logger.Info("loading inventory report", "path", inventoryPath)
	}

	wbInv, err := excelize.OpenFile(inventoryPath)
	if err != nil {
		return nil, fmt.Errorf("failed to open inventory report %s: %w", inventoryPath, err)
	}
	defer func() {
		_ = wbInv.Close()
	}()

	invRows, err := wbInv.GetRows("Sheet1")
	if err != nil {
		return nil, fmt.Errorf("failed to read inventory sheet: %w", err)
	}
	if len(invRows) < 1 {
		return nil, fmt.Errorf("inventory report appears empty")
	}

	inventoryBySKU := make(map[string]*inventoryEntry)
	for rowNum := 2; ; rowNum += 3 {
		item, stop := parseInventoryEntry(invRows, rowNum, logger)
		if stop {
			break
		}
		if item == nil {
			continue
		}
		inventoryBySKU[item.SKU] = item
	}

	return inventoryBySKU, nil
}

// parseInventoryEntry parses one inventory item from the worksheet rows and reports whether
// parsing should stop because the scan reached the workbook footer or ran out of rows.
func parseInventoryEntry(rows [][]string, rowNum int, logger *slog.Logger) (*inventoryEntry, bool) {
	if rowNum-1 >= len(rows) {
		return nil, true
	}

	// Item codes start at B2 and repeat every 3 rows in the inventory export.
	sku := getCellAt(rows, rowNum, inventorySKUIdx)
	if sku == "" {
		if logger != nil {
			logger.Info("Skipping empty SKU at inventory row", "row", rowNum)
		}
		return nil, false
	}
	// If the SKU cell looks like a run date, stop parsing so footer metadata is ignored.
	if isRunDate(sku) {
		if logger != nil {
			logger.Info("Encountered run-date/footer, stopping parse", "value", sku, "row", rowNum)
		}
		return nil, true
	}

	valRow := rowNum + 2
	if valRow-1 >= len(rows) {
		return nil, true
	}

	item := &inventoryEntry{SKU: sku}
	item.ProductLine = getCellAt(rows, valRow, inventoryProductLineIdx)
	item.ClassDesc = getCellAt(rows, valRow, inventoryClassIdx)
	item.RawClassDesc = item.ClassDesc
	item.Status = getCellAt(rows, valRow, inventoryStatusIdx)
	item.OnHand = parseInt(getCellAt(rows, valRow, inventoryOnHandIdx))
	item.OnPO = parseInt(getCellAt(rows, valRow, inventoryOnPOIdx))
	item.OnSO = parseInt(getCellAt(rows, valRow, inventoryOnSOIdx))
	item.OnBO = parseInt(getCellAt(rows, valRow, inventoryOnBOIdx))
	item.TotalAvailable = parseInt(getCellAt(rows, valRow, inventoryTotalAvailIdx))
	item.YTDSold = parseInt(getCellAt(rows, valRow, inventoryYTDSoldIdx))
	item.YTDIssued = parseInt(getCellAt(rows, valRow, inventoryYTDIssuedIdx))
	item.SoldPY = parseInt(getCellAt(rows, valRow, inventorySoldPYIdx))
	item.IssuedPY = parseInt(getCellAt(rows, valRow, inventoryIssuedPYIdx))
	item.Foil = getCellAt(rows, valRow, inventoryFoilIdx)
	item.Occasion = getCellAt(rows, valRow, inventoryOccasionIdx)
	item.Description = getCellAt(rows, valRow, inventoryDescIdx)
	item.UPC = getCellAt(rows, valRow, inventoryUPCIdx)
	item.RoyaltyCode = getCellAt(rows, valRow, inventoryRoyaltyCodeIdx)
	item.DollarSoldYTD = parseInventoryDollar(getCellAt(rows, valRow, inventoryDollarYTDIdx))
	item.DollarSoldPY = parseInventoryDollar(getCellAt(rows, valRow, inventoryDollarPYIdx))

	if logger != nil {
		logger.Debug("Inventory parse",
			"SKU", item.SKU,
			"skuRow", rowNum,
			"valRow", valRow,
			"ProductLine", item.ProductLine,
			"ClassDesc", item.ClassDesc,
			"Status", item.Status,
			"OnHand", item.OnHand,
			"OnPO", item.OnPO,
			"OnSO", item.OnSO,
			"OnBO", item.OnBO,
			"TotalAvailable", item.TotalAvailable,
			"YTDSold", item.YTDSold,
			"YTDIssued", item.YTDIssued,
			"SoldPY", item.SoldPY,
			"IssuedPY", item.IssuedPY,
			"Foil", item.Foil,
			"Occasion", item.Occasion,
			"Description", item.Description,
			"UPC", item.UPC,
			"RoyaltyCode", item.RoyaltyCode,
			"DollarSoldYTD", item.DollarSoldYTD,
			"DollarSoldPY", item.DollarSoldPY,
		)
	}

	return item, false
}

// parseInventoryDollar converts inventory currency text into a float64 while preserving the
// forgiving parsing used by the current workbook import.
func parseInventoryDollar(s string) float64 {
	trimmed := strings.TrimSpace(s)
	if trimmed == "" {
		return 0
	}
	trimmed = strings.ReplaceAll(trimmed, "$", "")
	trimmed = strings.ReplaceAll(trimmed, ",", "")
	trimmed = strings.ReplaceAll(trimmed, "(", "-")
	trimmed = strings.ReplaceAll(trimmed, ")", "")
	return parseFloat(trimmed)
}
