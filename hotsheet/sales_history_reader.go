package hotsheet

import (
	"errors"
	"fmt"
	"log/slog"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

const salesHistoryWarehouse = "BSC"

var errMissingSalesPeriod = errors.New("missing sales history period")

// mergeSalesHistory streams the Sage report at path into inventoryBySKU, using logger
// for unmatched items. Only BSC year blocks are attached; item and report totals are
// ignored. An empty path is a no-op; an unreadable or malformed report returns an error.
func mergeSalesHistory(path string, inventoryBySKU map[string]*inventoryEntry, logger *slog.Logger) error {
	if strings.TrimSpace(path) == "" {
		return nil
	}

	// Excelize's row iterator avoids holding this large, irregular worksheet in memory.
	workbook, err := excelize.OpenFile(path)
	if err != nil {
		return fmt.Errorf("failed to open sales history report %s: %w", path, err)
	}
	defer func() { _ = workbook.Close() }()

	rows, err := workbook.Rows("Sheet1")
	if err != nil {
		return fmt.Errorf("failed to read sales history sheet: %w", err)
	}
	defer func() { _ = rows.Close() }()

	var item *inventoryEntry
	var year int
	bscWarehouse := false
	rowNum := 0
	for rows.Next() {
		rowNum++
		// Raw values preserve the report's 0–100 percentage scale and numeric dollar values.
		cells, err := rows.Columns(excelize.Options{RawCellValue: true})
		if err != nil {
			return fmt.Errorf("failed to read sales history row %d: %w", rowNum, err)
		}
		label := strings.TrimSpace(getCell(cells, 0))
		if rowNum == 1 {
			if label != "Item Code" || getCell(cells, 1) != "Period 1" || getCell(cells, 12) != "Period 12" {
				return fmt.Errorf("sales history report has unexpected period headers")
			}
			continue
		}

		switch {
		case strings.TrimSpace(getCell(cells, 2)) == "Product Line:":
			item = inventoryBySKU[label]
			bscWarehouse = false
			year = 0
			if item == nil && logger != nil {
				logger.Debug("Skipping sales-history-only SKU", "SKU", label)
			}
		case label == "Warehouse:":
			// The warehouse code precedes the display name; matching just the code
			// avoids depending on a particular spelling of that display name.
			warehouse := strings.Fields(getCell(cells, 1))
			bscWarehouse = item != nil && len(warehouse) > 0 && warehouse[0] == salesHistoryWarehouse
			year = 0
		case strings.HasPrefix(label, "Total For Item:"):
			bscWarehouse = false
			year = 0
		case label == "Report Total:":
			return nil
		case label == "Year:":
			if !bscWarehouse {
				continue
			}
			year, err = strconv.Atoi(strings.TrimSpace(getCell(cells, 1)))
			if err != nil {
				return fmt.Errorf("invalid sales history year at row %d: %w", rowNum, err)
			}
		case isSalesHistoryMetric(label) && bscWarehouse:
			if year == 0 {
				return fmt.Errorf("sales history metric %q at row %d has no year", label, rowNum)
			}
			record, err := parseSalesRecord(cells, year, label)
			if err != nil {
				return fmt.Errorf("invalid sales history row %d: %w", rowNum, err)
			}
			item.SalesRecords = append(item.SalesRecords, record)
		}
	}
	if err := rows.Error(); err != nil {
		return fmt.Errorf("failed to scan sales history sheet: %w", err)
	}
	if rowNum == 0 {
		return fmt.Errorf("sales history report appears empty")
	}
	return nil
}

// isSalesHistoryMetric reports whether label names one of the five supported per-year metrics.
func isSalesHistoryMetric(label string) bool {
	switch label {
	case "Quantity Sold:", "Dollars Sold:", "Gross Profit Percent:", "Cost of Goods Sold:", "Quantity Returned:":
		return true
	default:
		return false
	}
}

// parseSalesRecord reads the twelve monthly values from cells into a year/metric record.
// It returns an error for missing or non-numeric periods rather than silently changing sales totals.
func parseSalesRecord(cells []string, year int, metric string) (salesRecord, error) {
	record := salesRecord{Year: year, Metric: strings.TrimSuffix(metric, ":")}
	for period := range record.Periods {
		value := strings.TrimSpace(getCell(cells, period+1))
		if value == "" {
			return salesRecord{}, fmt.Errorf("%w: period %d for %s", errMissingSalesPeriod, period+1, metric)
		}
		parsed, err := strconv.ParseFloat(strings.ReplaceAll(value, ",", ""), 64)
		if err != nil {
			return salesRecord{}, fmt.Errorf("invalid period %d for %s: %w", period+1, metric, err)
		}
		record.Periods[period] = parsed
	}
	return record, nil
}
