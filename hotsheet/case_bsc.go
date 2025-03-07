package hotsheet

import (
	"fmt"
)

// CaseBSC is a helper function for updating the BSC hotsheet. It creates an Update struct
// and calls its Update method.
//
// This function is called by the main function when the user selects 'BSC' as the product
// line.
func CaseBSC(fileHotsheetNew, inventoryReport, POReport string) error {
	everyday := Update{
		Hotsheet:          fileHotsheetNew,
		Sheet:             "Everyday",
		InventoryReport:   inventoryReport,
		POReport:          POReport,
		SkuCol:            "D",
		OnHandCol:         "E",
		OnPOCol1:          "F",
		OnPOCol2:          "H",
		OnPOCol3:          "J",
		OnPOColTotal:      "L",
		OnSOBOCol:         "M",
		YtdSoldIssuedCol:  "P",
		PONumCol1:         "G",
		PONumCol2:         "I",
		PONumCol3:         "K",
		AverageMonthlyCol: "Q",
	}
	winter := Update{
		Hotsheet:          fileHotsheetNew,
		Sheet:             "Winter Holiday",
		InventoryReport:   inventoryReport,
		POReport:          POReport,
		SkuCol:            "E",
		OnHandCol:         "F",
		OnPOCol1:          "I",
		OnPOCol2:          "K",
		OnPOCol3:          "M",
		OnPOColTotal:      "O",
		OnSOBOCol:         "G",
		YtdSoldIssuedCol:  "Q",
		PONumCol1:         "J",
		PONumCol2:         "L",
		PONumCol3:         "N",
		AverageMonthlyCol: "R",
	}
	spring := Update{
		Hotsheet:          fileHotsheetNew,
		Sheet:             "Spring Holiday",
		InventoryReport:   inventoryReport,
		POReport:          POReport,
		SkuCol:            "D",
		OnHandCol:         "E",
		OnPOCol1:          "G",
		OnPOCol2:          "I",
		OnPOCol3:          "K",
		OnPOColTotal:      "M",
		OnSOBOCol:         "N",
		YtdSoldIssuedCol:  "P",
		PONumCol1:         "H",
		PONumCol2:         "J",
		PONumCol3:         "L",
		AverageMonthlyCol: "Q",
	}
	a2Notecards := Update{
		Hotsheet:          fileHotsheetNew,
		Sheet:             "A2 Notecards",
		InventoryReport:   inventoryReport,
		POReport:          POReport,
		SkuCol:            "D",
		OnHandCol:         "E",
		OnPOCol1:          "F",
		OnPOCol2:          "H",
		OnPOCol3:          "J",
		OnPOColTotal:      "L",
		OnSOBOCol:         "M",
		YtdSoldIssuedCol:  "P",
		PONumCol1:         "G",
		PONumCol2:         "I",
		PONumCol3:         "K",
		AverageMonthlyCol: "Q",
	}

	// Update the hotsheet
	err := everyday.UpdateInventory("BSC", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update everyday stock: %w", err)
	}
	fmt.Println("Everyday product updated successfully")

	err = everyday.UpdatePONumber("BSC", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update everyday PO number: %w", err)
	}
	fmt.Println("Everyday PO number updated successfully")

	err = winter.UpdateInventory("BSC", "winter")
	if err != nil {
		return fmt.Errorf("failed to update holiday stock: %w", err)
	}
	fmt.Println("Holiday product updated successfully")

	err = winter.UpdatePONumber("BSC", "winter")
	if err != nil {
		return fmt.Errorf("failed to update holiday PO number: %w", err)
	}
	fmt.Println("Holiday PO number updated successfully")

	err = spring.UpdateInventory("BSC", "spring")
	if err != nil {
		return fmt.Errorf("failed to update spring stock: %w", err)
	}
	fmt.Println("Spring product updated successfully")

	err = spring.UpdatePONumber("BSC", "spring")
	if err != nil {
		return fmt.Errorf("failed to update spring PO number: %w", err)
	}
	fmt.Println("Spring PO number updated successfully")

	err = a2Notecards.UpdateInventory("BSC", "a2notecards")
	if err != nil {
		return fmt.Errorf("failed to update A2 Notecards stock: %w", err)
	}
	fmt.Println("A2 Notecard product updated successfully")

	err = a2Notecards.UpdatePONumber("BSC", "a2notecards")
	if err != nil {
		return fmt.Errorf("failed to update A2 Notecards PO number: %w", err)
	}
	fmt.Println("A2 Notecard PO number updated successfully")

	return nil
}
