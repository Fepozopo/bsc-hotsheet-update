package hotsheet

import (
	"fmt"
)

// CaseSMD is a helper function for updating the SMD hotsheet. It creates an Update struct
// and calls its Update method.
//
// This function is called by the main function when the user selects 'SMD' as the product
// line.
func CaseSMD(fileHotsheetNew, inventoryReport, POReport string) error {
	everyday := Update{
		Hotsheet:         fileHotsheetNew,
		Sheet:            "EVERYDAY",
		InventoryReport:  inventoryReport,
		POReport:         POReport,
		SkuCol:           "E",
		OnHandCol:        "F",
		OnPOCol1:         "G",
		OnPOCol2:         "I",
		OnPOCol3:         "K",
		OnPOColTotal:     "M",
		OnSOBOCol:        "N",
		YtdSoldIssuedCol: "S",
		PONumCol1:        "H",
		PONumCol2:        "J",
		PONumCol3:        "L",
	}
	holiday := Update{
		Hotsheet:         fileHotsheetNew,
		Sheet:            "HOLIDAY",
		InventoryReport:  inventoryReport,
		POReport:         POReport,
		SkuCol:           "D",
		OnHandCol:        "E",
		OnPOCol1:         "F",
		OnPOCol2:         "H",
		OnPOCol3:         "J",
		OnPOColTotal:     "L",
		OnSOBOCol:        "M",
		YtdSoldIssuedCol: "S",
		PONumCol1:        "G",
		PONumCol2:        "I",
		PONumCol3:        "K",
	}

	// Update the hotsheet
	err := everyday.UpdateInventory("SMD", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update everyday stock: %w", err)
	}
	fmt.Println("Everyday product updated successfully")

	err = everyday.UpdatePONumber("SMD", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update everyday PO number: %w", err)
	}
	fmt.Println("Everyday PO number updated successfully")

	err = holiday.UpdateInventory("SMD", "holiday")
	if err != nil {
		return fmt.Errorf("failed to update holiday stock: %w", err)
	}
	fmt.Println("Holiday product updated successfully")

	err = holiday.UpdatePONumber("SMD", "holiday")
	if err != nil {
		return fmt.Errorf("failed to update holiday PO number: %w", err)
	}
	fmt.Println("Holiday PO number updated successfully")

	return nil
}
