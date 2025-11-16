package hotsheet

import (
	"fmt"
)

// Case21C is a helper function for updating the 21C hotsheet. It creates an Update struct
// and calls its Update method.
//
// This function is called by the main function when the user selects '21C' as the product
// line.
func Case21C(fileHotsheetNew, inventoryReport, POReport, BNReport string) error {
	everyday := Update{
		Hotsheet:          fileHotsheetNew,
		Sheet:             "EVERYDAY",
		InventoryReport:   inventoryReport,
		POReport:          POReport,
		BNReport:          BNReport,
		SkuCol:            "C",
		OnHandCol:         "D",
		OnPOCol1:          "E",
		OnPOCol2:          "G",
		OnPOCol3:          "I",
		OnPOColTotal:      "K",
		OnSOBOCol:         "L",
		YtdSoldIssuedCol:  "Q",
		PONumCol1:         "F",
		PONumCol2:         "H",
		PONumCol3:         "J",
		AverageMonthlyCol: "R",
	}
	winter := Update{
		Hotsheet:          fileHotsheetNew,
		Sheet:             "Winter Holiday",
		InventoryReport:   inventoryReport,
		POReport:          POReport,
		BNReport:          BNReport,
		SkuCol:            "C",
		OnHandCol:         "D",
		OnPOCol1:          "E",
		OnPOCol2:          "G",
		OnPOCol3:          "I",
		OnPOColTotal:      "K",
		OnSOBOCol:         "L",
		YtdSoldIssuedCol:  "Q",
		PONumCol1:         "F",
		PONumCol2:         "H",
		PONumCol3:         "J",
		AverageMonthlyCol: "R",
	}
	spring := Update{
		Hotsheet:          fileHotsheetNew,
		Sheet:             "Spring Holiday",
		InventoryReport:   inventoryReport,
		POReport:          POReport,
		BNReport:          BNReport,
		SkuCol:            "C",
		OnHandCol:         "D",
		OnPOCol1:          "E",
		OnPOCol2:          "G",
		OnPOCol3:          "I",
		OnPOColTotal:      "K",
		OnSOBOCol:         "L",
		YtdSoldIssuedCol:  "Q",
		PONumCol1:         "F",
		PONumCol2:         "H",
		PONumCol3:         "J",
		AverageMonthlyCol: "R",
	}

	// Update the hotsheet
	err := everyday.UpdateInventory("21c", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update everyday stock: %w", err)
	}
	fmt.Println("Everyday product updated successfully")

	err = everyday.UpdatePONumber("21c", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update everyday PO number: %w", err)
	}
	fmt.Println("Everyday PO number updated successfully")

	err = winter.UpdateInventory("21c", "winter")
	if err != nil {
		return fmt.Errorf("failed to update stock: %w", err)
	}
	fmt.Println("Winter product updated successfully")

	err = winter.UpdatePONumber("21c", "winter")
	if err != nil {
		return fmt.Errorf("failed to update winter PO number: %w", err)
	}
	fmt.Println("Winter PO number updated successfully")

	err = spring.UpdateInventory("21c", "spring")
	if err != nil {
		return fmt.Errorf("failed to update stock: %w", err)
	}
	fmt.Println("Spring product updated successfully")

	err = spring.UpdatePONumber("21c", "spring")
	if err != nil {
		return fmt.Errorf("failed to update spring PO number: %w", err)
	}
	fmt.Println("Spring PO number updated successfully")

	return nil
}
