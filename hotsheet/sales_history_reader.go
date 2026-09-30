package hotsheet

import (
	"errors"
	"fmt"
	"log/slog"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const salesHistoryWarehouse = "BSC"

var (
	errMissingSalesPeriod         = errors.New("missing sales history period")
	errMissingSalesHistoryRunDate = errors.New("sales history report has no run date")
)

// mergeSalesHistory attaches BSC sales metrics to matching inventory SKUs and
// returns the Sage report date used to anchor MTO forecasts. Empty paths are ignored.
func mergeSalesHistory(path string, inventoryBySKU map[string]*inventoryEntry, logger *slog.Logger) (time.Time, error) {
	return mergeMonthlyHistory(path, inventoryBySKU, logger, false)
}

// mergeIssueHistory attaches BSC issued quantities to matching SKUs. It returns
// an error if the optional Sage issue export cannot be read or has no run date.
func mergeIssueHistory(path string, inventoryBySKU map[string]*inventoryEntry, logger *slog.Logger) error {
	_, err := mergeMonthlyHistory(path, inventoryBySKU, logger, true)
	return err
}

// mergeMonthlyHistory streams one Sage export, selecting only BSC item/year metrics
// (not warehouse, item or report totals). It returns its run date or an error;
// issueHistory selects issued units instead of the five sales-history metrics.
func mergeMonthlyHistory(path string, inventoryBySKU map[string]*inventoryEntry, logger *slog.Logger, issueHistory bool) (time.Time, error) {
	if strings.TrimSpace(path) == "" {
		return time.Time{}, nil
	}
	reportName := "sales history"
	if issueHistory {
		reportName = "issue history"
	}

	// Excelize's row iterator avoids holding this large, irregular worksheet in memory.
	workbook, err := excelize.OpenFile(path)
	if err != nil {
		return time.Time{}, fmt.Errorf("failed to open %s report %s: %w", reportName, path, err)
	}
	defer func() { _ = workbook.Close() }()

	rows, err := workbook.Rows("Sheet1")
	if err != nil {
		return time.Time{}, fmt.Errorf("failed to read %s sheet: %w", reportName, err)
	}
	defer func() { _ = rows.Close() }()

	var item *inventoryEntry
	var year int
	bscWarehouse := false
	var runDate time.Time
	rowNum := 0
	for rows.Next() {
		rowNum++
		// Raw values preserve the report's 0–100 percentage scale and numeric dollar values.
		cells, err := rows.Columns(excelize.Options{RawCellValue: true})
		if err != nil {
			return time.Time{}, fmt.Errorf("failed to read %s row %d: %w", reportName, rowNum, err)
		}
		label := strings.TrimSpace(getCell(cells, 0))
		if rowNum == 1 {
			if label != "Item Code" || getCell(cells, 1) != "Period 1" || getCell(cells, 12) != "Period 12" {
				return time.Time{}, fmt.Errorf("%s report has unexpected period headers", reportName)
			}
			continue
		}

		switch {
		case strings.TrimSpace(getCell(cells, 2)) == "Product Line:":
			item = inventoryBySKU[label]
			bscWarehouse = false
			year = 0
			if item == nil && logger != nil {
				logger.Debug("Skipping history-only SKU", "report", reportName, "SKU", label)
			}
		case label == "Warehouse:":
			// The warehouse code precedes the display name; matching just the code
			// avoids depending on a particular spelling of that display name.
			warehouse := strings.Fields(getCell(cells, 1))
			bscWarehouse = item != nil && len(warehouse) > 0 && warehouse[0] == salesHistoryWarehouse
			year = 0
		case strings.HasPrefix(label, "Total For Warehouse:"), strings.HasPrefix(label, "Total For Item:"):
			bscWarehouse = false
			year = 0
		case label == "Report Total:":
			// Report totals are not SKU sales. Continue to the run-date footer.
			item = nil
			bscWarehouse = false
			year = 0
		case strings.HasPrefix(label, "Run Date:"):
			fields := strings.Fields(strings.TrimPrefix(label, "Run Date:"))
			if len(fields) == 0 {
				return time.Time{}, fmt.Errorf("%s run date is empty at row %d", reportName, rowNum)
			}
			runDate, err = time.Parse("1/2/2006", fields[0])
			if err != nil {
				return time.Time{}, fmt.Errorf("invalid %s run date at row %d: %w", reportName, rowNum, err)
			}
		case label == "Year:":
			if !bscWarehouse {
				continue
			}
			year, err = strconv.Atoi(strings.TrimSpace(getCell(cells, 1)))
			if err != nil {
				return time.Time{}, fmt.Errorf("invalid %s year at row %d: %w", reportName, rowNum, err)
			}
		case bscWarehouse && ((issueHistory && label == "Quantity Issued:") || (!issueHistory && isSalesHistoryMetric(label))):
			if year == 0 {
				return time.Time{}, fmt.Errorf("%s metric %q at row %d has no year", reportName, label, rowNum)
			}
			record, err := parseSalesRecord(cells, year, label)
			if err != nil {
				return time.Time{}, fmt.Errorf("invalid %s row %d: %w", reportName, rowNum, err)
			}
			if issueHistory {
				// Sage may record returns as negative issues. They cannot offset
				// another month's shipments or a sale in the same month.
				for month := range record.Periods {
					record.Periods[month] = max(record.Periods[month], 0)
				}
			}
			item.SalesRecords = append(item.SalesRecords, record)
		}
	}
	if err := rows.Error(); err != nil {
		return time.Time{}, fmt.Errorf("failed to scan %s sheet: %w", reportName, err)
	}
	if rowNum == 0 {
		return time.Time{}, fmt.Errorf("%s report appears empty", reportName)
	}
	if runDate.IsZero() {
		if issueHistory {
			return time.Time{}, fmt.Errorf("issue history report has no run date")
		}
		return time.Time{}, errMissingSalesHistoryRunDate
	}
	return runDate, nil
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
