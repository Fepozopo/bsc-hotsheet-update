package hotsheet

import (
	"fmt"
)

// Case21C is a helper function for updating the 21C hotsheet. It creates an Update struct
// and calls its Update method.
//
// This function is called by the main function when the user selects '21C' as the product
// line.
func Case21C(fileHotsheetNew, inventoryReport, POReport string) error {
	everyday := Update{
		Hotsheet:         fileHotsheetNew,
		Sheet:            "EVERYDAY",
		InventoryReport:  inventoryReport,
		POReport:         POReport,
		SkuCol:           "C",
		OnHandCol:        "D",
		OnPOCol:          "E",
		OnSOBOCol:        "G",
		YtdSoldIssuedCol: "M",
		PONumCol:         "F",
	}
	winter := Update{
		Hotsheet:         fileHotsheetNew,
		Sheet:            "Winter Holiday",
		InventoryReport:  inventoryReport,
		POReport:         POReport,
		SkuCol:           "C",
		OnHandCol:        "D",
		OnPOCol:          "E",
		OnSOBOCol:        "G",
		YtdSoldIssuedCol: "M",
		PONumCol:         "F",
	}
	spring := Update{
		Hotsheet:         fileHotsheetNew,
		Sheet:            "Spring Holiday",
		InventoryReport:  inventoryReport,
		POReport:         POReport,
		SkuCol:           "C",
		OnHandCol:        "D",
		OnPOCol:          "E",
		OnSOBOCol:        "G",
		YtdSoldIssuedCol: "M",
		PONumCol:         "F",
	}

	// Update the hotsheet
	err := everyday.Update("21c", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update everyday stock: %w", err)
	}
	fmt.Println("Everyday product updated successfully")

	err = everyday.UpdatePONumber("21c", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update everyday PO number: %w", err)
	}
	fmt.Println("Everyday PO number updated successfully")

	err = winter.Update("21c", "winter")
	if err != nil {
		return fmt.Errorf("failed to update stock: %w", err)
	}
	fmt.Println("Winter product updated successfully")

	err = winter.UpdatePONumber("21c", "winter")
	if err != nil {
		return fmt.Errorf("failed to update winter PO number: %w", err)
	}
	fmt.Println("Winter PO number updated successfully")

	err = spring.Update("21c", "spring")
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
