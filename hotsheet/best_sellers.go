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

// BestSellersRange selects an inclusive pair of calendar months from the BSC
// history report. A nil range instead selects the required inventory report's YTD sales.
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
		return fmt.Errorf("Best Sellers range needs valid years and months")
	}
	if r.FromYear*12+r.FromMonth > r.ToYear*12+r.ToMonth {
		return fmt.Errorf("Best Sellers from month must not be after to month")
	}
	return nil
}

// bestSellerRow holds the source inventory item and its sales for the requested period.
type bestSellerRow struct {
	item     *inventoryEntry
	quantity float64
	dollars  float64
}

// bestSellerRows calculates sales for entries and returns them ranked by quantity
// sold descending, breaking ties by SKU. A nil range uses inventory YTD sales.
func bestSellerRows(entries []*inventoryEntry, period *BestSellersRange) []bestSellerRow {
	rows := make([]bestSellerRow, 0, len(entries))
	for _, item := range entries {
		row := bestSellerRow{item: item}
		if period == nil {
			row.quantity = float64(item.YTDSold)
			row.dollars = item.DollarSoldYTD
		} else {
			for _, record := range item.SalesRecords {
				if record.Metric != "Quantity Sold" && record.Metric != "Dollars Sold" {
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
					if record.Metric == "Quantity Sold" {
						row.quantity += record.Periods[month-1]
					} else {
						row.dollars += record.Periods[month-1]
					}
				}
			}
		}
		rows = append(rows, row)
	}
	sort.Slice(rows, func(i, j int) bool {
		if rows[i].quantity != rows[j].quantity {
			return rows[i].quantity > rows[j].quantity
		}
		return rows[i].item.SKU < rows[j].item.SKU
	})
	return rows
}

// writeBestSellersSheet adds one ranked inventory row per SKU to f. If period is
// nil it writes inventory YTD sales; otherwise it sums the inclusive BSC months.
// It returns an error when a worksheet operation fails.
func writeBestSellersSheet(f *excelize.File, entries []*inventoryEntry, period *BestSellersRange) error {
	if _, err := f.NewSheet(bestSellersSheetName); err != nil {
		return fmt.Errorf("failed to create Best Sellers sheet: %w", err)
	}
	headers := [...]string{"Item Code", "Description", "Quantity Sold", "Dollars Sold", "Royalty Code", "Class Description", "Occasion", "Foil Status", "Status"}
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
		case "Quantity Sold", "Dollars Sold":
			width = 18
		case "Class Description", "Foil Status":
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
		values := [...]interface{}{item.SKU, item.Description, row.quantity, row.dollars, item.RoyaltyCode, classDesc, item.Occasion, item.Foil, item.Status}
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
		if err := f.SetCellStyle(bestSellersSheetName, fmt.Sprintf("A%d", rowNum), fmt.Sprintf("I%d", rowNum), rowStyles[statusIndex]); err != nil {
			return fmt.Errorf("failed to style Best Sellers row %d: %w", rowNum, err)
		}
		if err := f.SetCellStyle(bestSellersSheetName, fmt.Sprintf("D%d", rowNum), fmt.Sprintf("D%d", rowNum), dollarStyles[statusIndex]); err != nil {
			return fmt.Errorf("failed to format Best Sellers dollars at row %d: %w", rowNum, err)
		}
	}
	// Excelize writes one filter dropdown per header, including the final Status field.
	if err := f.AutoFilter(bestSellersSheetName, "A1:I1", nil); err != nil {
		return fmt.Errorf("failed to set Best Sellers autofilter: %w", err)
	}
	return nil
}
