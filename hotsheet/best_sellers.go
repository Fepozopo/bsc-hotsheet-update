package hotsheet

import (
	"errors"
	"fmt"
	"sort"

	"github.com/xuri/excelize/v2"
)

const bestSellersSheetName = "Best Sellers"

// ErrBestSellersHistoryRequired indicates that a month range was requested without
// the optional sales-history report needed to supply its monthly sales figures.
var ErrBestSellersHistoryRequired = errors.New("a sales history report is required for a Best Sellers date range")

// BestSellersRange selects inclusive BSC sales and optional issue-history months.
// A nil range instead selects the required inventory report's YTD shipments.
type BestSellersRange struct {
	FromYear  int
	FromMonth int
	ToYear    int
	ToMonth   int
}

// Validate reports whether both endpoints are real months and the start is no
// later than the end. Callers should also require a history file for a range.
func (r BestSellersRange) Validate() error {
	if r.FromYear < 1 || r.FromYear > 9999 || r.ToYear < 1 || r.ToYear > 9999 || r.FromMonth < 1 || r.FromMonth > 12 || r.ToMonth < 1 || r.ToMonth > 12 {
		return fmt.Errorf("best sellers range needs valid years and months")
	}
	if r.FromYear*12+r.FromMonth > r.ToYear*12+r.ToMonth {
		return fmt.Errorf("best sellers from month must not be after to month")
	}
	return nil
}

// bestSellerRow holds an inventory item, shipped units and sales dollars for
// the requested period. Shipped units combine sold and nonnegative issued units.
type bestSellerRow struct {
	item    *inventoryEntry
	shipped float64
	dollars float64
}

// bestSellerRows ranks entries by monthly BSC sold plus nonnegative issued units,
// breaking ties by SKU. A nil range uses the inventory YTD sold and issued
// snapshot and YTD sales dollars, not the optional monthly history.
func bestSellerRows(entries []*inventoryEntry, period *BestSellersRange) []bestSellerRow {
	rows := make([]bestSellerRow, 0, len(entries))
	for _, item := range entries {
		row := bestSellerRow{item: item}
		if period == nil {
			row.shipped = float64(item.YTDSold + max(item.YTDIssued, 0))
			row.dollars = item.DollarSoldYTD
		} else {
			for _, record := range item.SalesRecords {
				if record.Metric != "Quantity Sold" && record.Metric != "Quantity Issued" && record.Metric != "Dollars Sold" {
					continue
				}
				if record.Year < period.FromYear || record.Year > period.ToYear {
					continue
				}
				first, last := 1, 12
				if record.Year == period.FromYear {
					first = period.FromMonth
				}
				if record.Year == period.ToYear {
					last = period.ToMonth
				}
				for month := first; month <= last; month++ {
					switch record.Metric {
					case "Quantity Sold":
						row.shipped += record.Periods[month-1]
					case "Quantity Issued":
						row.shipped += max(record.Periods[month-1], 0)
					case "Dollars Sold":
						row.dollars += record.Periods[month-1]
					}
				}
			}
		}
		rows = append(rows, row)
	}
	sort.Slice(rows, func(i, j int) bool {
		if rows[i].shipped != rows[j].shipped {
			return rows[i].shipped > rows[j].shipped
		}
		return rows[i].item.SKU < rows[j].item.SKU
	})
	return rows
}

// writeBestSellersSheet adds one ranked row per SKU to f. A nil period uses
// inventory YTD shipped units and sales dollars; a range sums BSC sold plus
// issued units and sales dollars across inclusive months. Inventory quantities
// remain the current snapshot; committed units include sales orders and back
// orders. Worksheet failures are returned.
func writeBestSellersSheet(f *excelize.File, entries []*inventoryEntry, period *BestSellersRange) error {
	if _, err := f.NewSheet(bestSellersSheetName); err != nil {
		return fmt.Errorf("failed to create Best Sellers sheet: %w", err)
	}
	headers := [...]string{"Item Code", "Description", "Quantity Shipped", "Dollars Sold", "Quantity on Hand", "Quantity Committed", "Quantity on PO", "Royalty Code", "Class Description", "Occasion", "Foil Status", "Status"}
	headerStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(), Border: thinBlackBorder(),
		Fill: patternFill(standardHeaderFill), Font: boldFont(),
	})
	if err != nil {
		return fmt.Errorf("failed to create Best Sellers header style: %w", err)
	}
	for col, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		if err := f.SetCellValue(bestSellersSheetName, cell, header); err != nil {
			return fmt.Errorf("failed to write Best Sellers header %s: %w", cell, err)
		}
		if err := f.SetCellStyle(bestSellersSheetName, cell, cell, headerStyle); err != nil {
			return fmt.Errorf("failed to style Best Sellers header %s: %w", cell, err)
		}
		width := standardSheetWidthForHeader(header)
		switch header {
		case "Quantity Shipped", "Dollars Sold":
			width = 18
		case "Quantity on Hand", "Quantity Committed", "Quantity on PO", "Class Description", "Foil Status":
			width = 22
		}
		name, _ := excelize.ColumnNumberToName(col + 1)
		if err := f.SetColWidth(bestSellersSheetName, name, name, width); err != nil {
			return fmt.Errorf("failed to set Best Sellers column %s width: %w", name, err)
		}
	}

	var rowStyles, dollarStyles [3]int
	for i, status := range []string{"Carryover", "Rundown", "Discontinued"} {
		fill := standardSheetCellFillColor(status, 0, -1, -1, 0, 0, nil)
		base := &excelize.Style{Alignment: centeredAlignment(), Border: thinBlackBorder(), Fill: patternFill(fill)}
		rowStyles[i], err = f.NewStyle(base)
		if err != nil {
			return fmt.Errorf("failed to create Best Sellers row style: %w", err)
		}
		currency := *base
		currency.CustomNumFmt = currencyNumFmt()
		dollarStyles[i], err = f.NewStyle(&currency)
		if err != nil {
			return fmt.Errorf("failed to create Best Sellers currency style: %w", err)
		}
	}

	for index, row := range bestSellerRows(entries, period) {
		item := row.item
		rowNum := index + 2
		// RawClassDesc retains the inventory classification if the standard sheet
		// already added an SKU-derived display prefix to ClassDesc.
		classDesc := item.RawClassDesc
		if classDesc == "" {
			classDesc = item.ClassDesc
		}
		committed := item.OnSO + item.OnBO
		values := [...]interface{}{item.SKU, item.Description, row.shipped, row.dollars, item.OnHand, committed, item.OnPO, item.RoyaltyCode, classDesc, item.Occasion, item.Foil, item.Status}
		for col, value := range values {
			cell, _ := excelize.CoordinatesToCellName(col+1, rowNum)
			if err := f.SetCellValue(bestSellersSheetName, cell, value); err != nil {
				return fmt.Errorf("failed to write Best Sellers cell %s: %w", cell, err)
			}
		}
		statusIndex := 0
		switch item.Status {
		case "Rundown":
			statusIndex = 1
		case "Discontinued":
			statusIndex = 2
		}
		if err := f.SetCellStyle(bestSellersSheetName, fmt.Sprintf("A%d", rowNum), fmt.Sprintf("L%d", rowNum), rowStyles[statusIndex]); err != nil {
			return fmt.Errorf("failed to style Best Sellers row %d: %w", rowNum, err)
		}
		if err := f.SetCellStyle(bestSellersSheetName, fmt.Sprintf("D%d", rowNum), fmt.Sprintf("D%d", rowNum), dollarStyles[statusIndex]); err != nil {
			return fmt.Errorf("failed to format Best Sellers dollars at row %d: %w", rowNum, err)
		}
	}
	// Excelize writes one filter dropdown per header, including the final Status field.
	if err := f.AutoFilter(bestSellersSheetName, "A1:L1", nil); err != nil {
		return fmt.Errorf("failed to set Best Sellers autofilter: %w", err)
	}
	return nil
}
