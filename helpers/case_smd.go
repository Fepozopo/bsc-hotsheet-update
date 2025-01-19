package helpers

import (
	"fmt"
)

func CaseSMD() error {
	fileHotsheetNew, fileStockReport, fileSalesReport, err := GetFiles("SMD")
	if err != nil {
		return fmt.Errorf("failed to get files: %w", err)
	}

	// HOTSHEET | SECTION | REPORT | SKU | ON HAND | ON PO | ON SO/BO
	stock := UpdateStock{fileHotsheetNew, "EVERYDAY", fileStockReport, "E", "F", "I", "K"}
	stockHoliday := UpdateStock{fileHotsheetNew, "HOLIDAY", fileStockReport, "C", "D", "F", "H"}
	// HOTSHEET | SECTION | REPORT | SKU | YTD
	sales := UpdateSales{fileHotsheetNew, "EVERYDAY", fileSalesReport, "E", "P"}
	salesHoliday := UpdateSales{fileHotsheetNew, "HOLIDAY", fileSalesReport, "C", "N"}

	// Update the hotsheet
	err = stock.UpdateStock()
	if err != nil {
		return fmt.Errorf("failed to update stock: %w", err)
	}
	fmt.Println("Stock updated successfully")

	err = stockHoliday.UpdateStock()
	if err != nil {
		return fmt.Errorf("failed to update holiday stock: %w", err)
	}
	fmt.Println("Holiday stock updated successfully")

	err = sales.UpdateSales()
	if err != nil {
		return fmt.Errorf("failed to update sales: %w", err)
	}
	fmt.Println("Sales updated successfully")

	err = salesHoliday.UpdateSales()
	if err != nil {
		return fmt.Errorf("failed to update holiday sales: %w", err)
	}
	fmt.Println("Holiday sales updated successfully")

	return nil
}
