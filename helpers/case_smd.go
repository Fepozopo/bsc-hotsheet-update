package helpers

import (
	"fmt"
)

func CaseSMD(fileHotsheetNew, fileStockReport, fileSalesReport string) error {
	// HOTSHEET | SECTION | REPORT | SKU | ON HAND | ON PO | ON SO/BO
	stock := UpdateStock{fileHotsheetNew, "EVERYDAY", fileStockReport, "E", "F", "I", "K"}
	stockHoliday := UpdateStock{fileHotsheetNew, "HOLIDAY", fileStockReport, "C", "D", "F", "H"}
	// HOTSHEET | SECTION | REPORT | SKU | YTD
	sales := UpdateSales{fileHotsheetNew, "EVERYDAY", fileSalesReport, "E", "P"}
	salesHoliday := UpdateSales{fileHotsheetNew, "HOLIDAY", fileSalesReport, "C", "N"}

	// Update the hotsheet
	fmt.Println("Updating stock...")
	err := stock.UpdateStock()
	if err != nil {
		return fmt.Errorf("failed to update stock: %w", err)
	}

	fmt.Println("Updating holiday stock...")
	err = stockHoliday.UpdateStock()
	if err != nil {
		return fmt.Errorf("failed to update holiday stock: %w", err)
	}

	fmt.Println("Updating sales...")
	err = sales.UpdateSales()
	if err != nil {
		return fmt.Errorf("failed to update sales: %w", err)
	}

	fmt.Println("Updating holiday sales...")
	err = salesHoliday.UpdateSales()
	if err != nil {
		return fmt.Errorf("failed to update holiday sales: %w", err)
	}

	return nil
}
