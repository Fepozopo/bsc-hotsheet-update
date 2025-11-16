package hotsheet

import (
	"fmt"
)

// CaseBJP is a helper function for updating the BJP hotsheet. It creates an Update struct
// and calls its Update method.
//
// This function is called by the main function when the user selects 'BJP' as the product
// line.
func CaseBJP(fileHotsheetNew, inventoryReport, POReport, BNReport string) error {
	everyday := Update{
		Hotsheet:            fileHotsheetNew,
		Sheet:               "Everyday",
		InventoryReport:     inventoryReport,
		POReport:            POReport,
		BNReport:            BNReport,
		SkuCol:              "C",
		OnHandCol:           "D",
		OnPOCol1:            "E",
		OnPOCol2:            "G",
		OnPOCol3:            "I",
		OnPOColTotal:        "K",
		OnSOBOCol:           "L",
		YtdSoldIssuedCol:    "N",
		PONumCol1:           "F",
		PONumCol2:           "H",
		PONumCol3:           "J",
		AverageMonthlyCol:   "O",
		BNYtdSoldCol:        "Q",
		BNAverageMonthlyCol: "R",
	}

	// Update the hotsheet
	err := everyday.UpdateInventory("BJP", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update everyday stock: %w", err)
	}
	fmt.Println("Everyday product updated successfully")

	err = everyday.UpdatePONumber("BJP", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update everyday PO number: %w", err)
	}
	fmt.Println("Everyday PO number updated successfully")

	return nil
}
