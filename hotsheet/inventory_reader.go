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
func loadInventoryEntries(inventoryPath string, logger *slog.Logger) (map[string]*entry, error) {
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

	invMap := make(map[string]*entry)
	for rowNum := 2; ; rowNum += 3 {
		e, stop := parseInventoryEntry(invRows, rowNum, logger)
		if stop {
			break
		}
		if e == nil {
			continue
		}
		invMap[e.SKU] = e
	}

	return invMap, nil
}

// parseInventoryEntry parses one inventory item from the worksheet rows and reports whether
// parsing should stop because the scan reached the workbook footer or ran out of rows.
func parseInventoryEntry(rows [][]string, rowNum int, logger *slog.Logger) (*entry, bool) {
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

	e := &entry{SKU: sku}
	e.ProductLine = getCellAt(rows, valRow, inventoryProductLineIdx)
	e.ClassDesc = getCellAt(rows, valRow, inventoryClassIdx)
	e.RawClassDesc = e.ClassDesc
	e.Status = getCellAt(rows, valRow, inventoryStatusIdx)
	e.OnHand = parseInt(getCellAt(rows, valRow, inventoryOnHandIdx))
	e.OnPO = parseInt(getCellAt(rows, valRow, inventoryOnPOIdx))
	e.OnSO = parseInt(getCellAt(rows, valRow, inventoryOnSOIdx))
	e.OnBO = parseInt(getCellAt(rows, valRow, inventoryOnBOIdx))
	e.TotalAvailable = parseInt(getCellAt(rows, valRow, inventoryTotalAvailIdx))
	e.YTDSold = parseInt(getCellAt(rows, valRow, inventoryYTDSoldIdx))
	e.YTDIssued = parseInt(getCellAt(rows, valRow, inventoryYTDIssuedIdx))
	e.SoldPY = parseInt(getCellAt(rows, valRow, inventorySoldPYIdx))
	e.IssuedPY = parseInt(getCellAt(rows, valRow, inventoryIssuedPYIdx))
	e.Foil = getCellAt(rows, valRow, inventoryFoilIdx)
	e.Occasion = getCellAt(rows, valRow, inventoryOccasionIdx)
	e.Description = getCellAt(rows, valRow, inventoryDescIdx)
	e.UPC = getCellAt(rows, valRow, inventoryUPCIdx)
	e.RoyaltyCode = getCellAt(rows, valRow, inventoryRoyaltyCodeIdx)
	e.DollarSoldYTD = parseInventoryDollar(getCellAt(rows, valRow, inventoryDollarYTDIdx))
	e.DollarSoldPY = parseInventoryDollar(getCellAt(rows, valRow, inventoryDollarPYIdx))

	if logger != nil {
		logger.Debug("Inventory parse",
			"SKU", e.SKU,
			"skuRow", rowNum,
			"valRow", valRow,
			"ProductLine", e.ProductLine,
			"ClassDesc", e.ClassDesc,
			"Status", e.Status,
			"OnHand", e.OnHand,
			"OnPO", e.OnPO,
			"OnSO", e.OnSO,
			"OnBO", e.OnBO,
			"TotalAvailable", e.TotalAvailable,
			"YTDSold", e.YTDSold,
			"YTDIssued", e.YTDIssued,
			"SoldPY", e.SoldPY,
			"IssuedPY", e.IssuedPY,
			"Foil", e.Foil,
			"Occasion", e.Occasion,
			"Description", e.Description,
			"UPC", e.UPC,
			"RoyaltyCode", e.RoyaltyCode,
			"DollarSoldYTD", e.DollarSoldYTD,
			"DollarSoldPY", e.DollarSoldPY,
		)
	}

	return e, false
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
