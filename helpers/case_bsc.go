package helpers

import (
	"fmt"
)

func CaseBSC(fileHotsheetNew, fileStockReport, fileSalesReport string) error {
	// HOTSHEET | SHEET | REPORT | SKU | ON HAND | ON PO | ON SO/BO
	stockEveryday := UpdateStock{fileHotsheetNew, "Everyday", fileStockReport, "D", "E", "F", "H"}
	stockWinter := UpdateStock{fileHotsheetNew, "Winter Holiday", fileStockReport, "E", "F", "I", "G"}
	stockSpring := UpdateStock{fileHotsheetNew, "Spring holiday", fileStockReport, "D", "E", "H", "J"}
	stockA2Notecards := UpdateStock{fileHotsheetNew, "A2 Notecards", fileStockReport, "D", "F", "G", "I"}
	// HOTSHEET | SHEET | REPORT | SKU | YTD
	salesEveryday := UpdateSales{fileHotsheetNew, "Everyday", fileSalesReport, "D", "K"}
	salesWinter := UpdateSales{fileHotsheetNew, "Winter Holiday", fileSalesReport, "E", "L"}
	salesWinterKit := UpdateSales{fileHotsheetNew, "Winter Holiday Kits", fileSalesReport, "E", "L"}
	salesSpring := UpdateSales{fileHotsheetNew, "Spring holiday", fileSalesReport, "D", "L"}
	salesA2Notecards := UpdateSales{fileHotsheetNew, "A2 Notecards", fileSalesReport, "D", "L"}

	// Update the hotsheet
	err := stockEveryday.UpdateStock("bsc", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update stock: %w", err)
	}
	fmt.Println("Everyday stock updated successfully")

	err = stockWinter.UpdateStock("bsc", "winter")
	if err != nil {
		return fmt.Errorf("failed to update holiday stock: %w", err)
	}
	fmt.Println("Holiday stock updated successfully")

	err = stockSpring.UpdateStock("bsc", "spring")
	if err != nil {
		return fmt.Errorf("failed to update spring stock: %w", err)
	}
	fmt.Println("Spring stock updated successfully")

	err = stockA2Notecards.UpdateStock("bsc", "a2notecards")
	if err != nil {
		return fmt.Errorf("failed to update A2 Notecards stock: %w", err)
	}
	fmt.Println("A2 Notecards stock updated successfully")

	err = salesEveryday.UpdateSales("bsc", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update sales: %w", err)
	}
	fmt.Println("Everyday sales updated successfully")

	err = salesWinter.UpdateSales("bsc", "winter")
	if err != nil {
		return fmt.Errorf("failed to update holiday sales: %w", err)
	}
	fmt.Println("Holiday sales updated successfully")

	err = salesWinterKit.UpdateSales("bsc", "winterkit")
	if err != nil {
		return fmt.Errorf("failed to update holiday kit sales: %w", err)
	}
	fmt.Println("Holiday kit sales updated successfully")

	err = salesSpring.UpdateSales("bsc", "spring")
	if err != nil {
		return fmt.Errorf("failed to update spring sales: %w", err)
	}
	fmt.Println("Spring sales updated successfully")

	err = salesA2Notecards.UpdateSales("bsc", "a2notecards")
	if err != nil {
		return fmt.Errorf("failed to update A2 Notecards sales: %w", err)
	}
	fmt.Println("A2 Notecards sales updated successfully")

	return nil
}
