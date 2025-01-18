package helpers

import (
	"fmt"
)

func Case21C() error {
	fileHotsheetNew, fileStockReport, fileSalesReport, err := GetFiles("2021co")
	if err != nil {
		return fmt.Errorf("failed to get files: %w", err)
	}

	// HOTSHEET | SECTION | REPORT | SKU | ON HAND | ON PO | ON SO/BO
	stock := UpdateStock{fileHotsheetNew, "EVERYDAY", fileStockReport, "C", "D", "E", "G"}
	// HOTSHEET | SECTION | REPORT | SKU | YTD
	sales := UpdateSales{fileHotsheetNew, "EVERYDAY", fileSalesReport, "C", "M"}
	salesKits := UpdateSales{fileHotsheetNew, "boxed card unit sales", fileSalesReport, "C", "H"}

	// Update the hotsheet
	err = stock.UpdateStock()
	if err != nil {
		return fmt.Errorf("failed to update stock: %w", err)
	}
	fmt.Println("Stock updated successfully")

	err = sales.UpdateSales()
	if err != nil {
		return fmt.Errorf("failed to update sales: %w", err)
	}
	fmt.Println("Sales updated successfully")

	err = salesKits.UpdateSales()
	if err != nil {
		return fmt.Errorf("failed to update kit sales: %w", err)
	}
	fmt.Println("Kit sales updated successfully")

	return nil
}
