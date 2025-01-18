package main

import (
	"fmt"
)

func handlerCaseBSC() error {
	fileHotsheetNew, fileStockReport, fileSalesReport, err := handlerGetFiles("BSC")
	if err != nil {
		return fmt.Errorf("failed to get files: %w", err)
	}

	// HOTSHEET | SECTION | REPORT | SKU | ON HAND | ON PO | ON SO/BO
	stock := UpdateStock{fileHotsheetNew, "Everyday", fileStockReport, "D", "E", "F", "H"}
	stockHoliday := UpdateStock{fileHotsheetNew, "Winter Holiday", fileStockReport, "E", "F", "I", "G"}
	stockSpring := UpdateStock{fileHotsheetNew, "Spring holiday", fileStockReport, "D", "E", "H", "J"}
	stockA2Notecards := UpdateStock{fileHotsheetNew, "A2 Notecards", fileStockReport, "D", "F", "G", "I"}
	// HOTSHEET | SECTION | REPORT | SKU | YTD
	sales := UpdateSales{fileHotsheetNew, "Everyday", fileSalesReport, "D", "K"}
	salesHoliday := UpdateSales{fileHotsheetNew, "Winter Holiday", fileSalesReport, "E", "L"}
	salesHolidayKit := UpdateSales{fileHotsheetNew, "Winter Holiday Kits", fileSalesReport, "E", "L"}
	salesSpring := UpdateSales{fileHotsheetNew, "Spring holiday", fileSalesReport, "D", "L"}
	salesA2Notecards := UpdateSales{fileHotsheetNew, "A2 Notecards", fileSalesReport, "D", "L"}

	// Update the hotsheet
	err = stock.handlerUpdateStock()
	if err != nil {
		return fmt.Errorf("failed to update stock: %w", err)
	}
	fmt.Println("Stock updated successfully")

	err = stockHoliday.handlerUpdateStock()
	if err != nil {
		return fmt.Errorf("failed to update holiday stock: %w", err)
	}
	fmt.Println("Holiday stock updated successfully")

	err = stockSpring.handlerUpdateStock()
	if err != nil {
		return fmt.Errorf("failed to update spring stock: %w", err)
	}
	fmt.Println("Spring stock updated successfully")

	err = stockA2Notecards.handlerUpdateStock()
	if err != nil {
		return fmt.Errorf("failed to update A2 Notecards stock: %w", err)
	}
	fmt.Println("A2 Notecards stock updated successfully")

	err = sales.handlerUpdateSales()
	if err != nil {
		return fmt.Errorf("failed to update sales: %w", err)
	}
	fmt.Println("Sales updated successfully")

	err = salesHoliday.handlerUpdateSales()
	if err != nil {
		return fmt.Errorf("failed to update holiday sales: %w", err)
	}
	fmt.Println("Holiday sales updated successfully")

	err = salesHolidayKit.handlerUpdateSales()
	if err != nil {
		return fmt.Errorf("failed to update holiday kit sales: %w", err)
	}
	fmt.Println("Holiday kit sales updated successfully")

	err = salesSpring.handlerUpdateSales()
	if err != nil {
		return fmt.Errorf("failed to update spring sales: %w", err)
	}
	fmt.Println("Spring sales updated successfully")

	err = salesA2Notecards.handlerUpdateSales()
	if err != nil {
		return fmt.Errorf("failed to update A2 Notecards sales: %w", err)
	}
	fmt.Println("A2 Notecards sales updated successfully")

	return nil
}
