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
		OnPOCol:          "I",
		OnSOBOCol:        "K",
		YtdSoldIssuedCol: "P",
		PONumCol:         "J",
	}
	holiday := Update{
		Hotsheet:         fileHotsheetNew,
		Sheet:            "HOLIDAY",
		InventoryReport:  inventoryReport,
		POReport:         POReport,
		SkuCol:           "C",
		OnHandCol:        "D",
		OnPOCol:          "F",
		OnSOBOCol:        "H",
		YtdSoldIssuedCol: "N",
		PONumCol:         "G",
	}

	// Update the hotsheet
	err := everyday.Update("SMD", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update everyday stock: %w", err)
	}
	fmt.Println("Everyday product updated successfully")

	err = everyday.UpdatePONumber("SMD", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update everyday PO number: %w", err)
	}
	fmt.Println("Everyday PO number updated successfully")

	err = holiday.Update("SMD", "holiday")
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
