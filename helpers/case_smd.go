package helpers

import (
	"fmt"
)

func CaseSMD(fileHotsheetNew, fileStockReport, fileSalesReport string) error {
	// HOTSHEET | SHEET | REPORT | SKU | ON HAND | ON PO | ON SO/BO
	stockEveryday := UpdateStock{fileHotsheetNew, "EVERYDAY", fileStockReport, "E", "F", "I", "K"}
	stockHoliday := UpdateStock{fileHotsheetNew, "HOLIDAY", fileStockReport, "C", "D", "F", "H"}
	// HOTSHEET | SHEET | REPORT | SKU | YTD
	salesEveryday := UpdateSales{fileHotsheetNew, "EVERYDAY", fileSalesReport, "E", "P"}
	salesHoliday := UpdateSales{fileHotsheetNew, "HOLIDAY", fileSalesReport, "C", "N"}

	// Update the hotsheet
	err := stockEveryday.UpdateStock("smd", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update stock: %w", err)
	}
	fmt.Println("Everyday stock updated successfully")

	err = stockHoliday.UpdateStock("smd", "holiday")
	if err != nil {
		return fmt.Errorf("failed to update holiday stock: %w", err)
	}
	fmt.Println("Holiday stock updated successfully")

	err = salesEveryday.UpdateSales("smd", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update sales: %w", err)
	}
	fmt.Println("Everyday sales updated successfully")

	err = salesHoliday.UpdateSales("smd", "holiday")
	if err != nil {
		return fmt.Errorf("failed to update holiday sales: %w", err)
	}
	fmt.Println("Holiday sales updated successfully")

	return nil
}
