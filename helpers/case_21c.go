package helpers

import (
	"fmt"
)

// Case21C is a helper function for updating the 21C hotsheet. It creates UpdateStock and UpdateSales structs
// and calls their respective UpdateStock and UpdateSales methods.
//
// This function is called by the main function when the user selects '21C' as the product
// line.
func Case21C(fileHotsheetNew, fileStockReport, fileSalesReport string) error {
	// HOTSHEET | SHEET | REPORT | SKU | ON HAND | ON PO | ON SO/BO
	stockEveryday := UpdateStock{fileHotsheetNew, "EVERYDAY", fileStockReport, "C", "D", "E", "G"}
	// HOTSHEET | SHEET | REPORT | SKU | YTD
	salesEveryday := UpdateSales{fileHotsheetNew, "EVERYDAY", fileSalesReport, "C", "M"}
	salesKits := UpdateSales{fileHotsheetNew, "boxed card unit sales", fileSalesReport, "C", "H"}

	// Update the hotsheet
	err := stockEveryday.UpdateStock("21c", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update stock: %w", err)
	}
	fmt.Println("Everyday stock updated successfully")

	err = salesEveryday.UpdateSales("21c", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update sales: %w", err)
	}
	fmt.Println("Sales updated successfully")

	err = salesKits.UpdateSales("21c", "kits")
	if err != nil {
		return fmt.Errorf("failed to update kit sales: %w", err)
	}
	fmt.Println("Kit sales updated successfully")

	return nil
}
