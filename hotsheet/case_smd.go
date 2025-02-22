package hotsheet

import (
	"fmt"
)

// CaseSMD is a helper function for updating the SMD hotsheet. It creates an Update struct
// and calls its Update method.
//
// This function is called by the main function when the user selects 'SMD' as the product
// line.
func CaseSMD(fileHotsheetNew, fileReport string) error {
	everyday := Update{
		Hotsheet:         fileHotsheetNew,
		Sheet:            "EVERYDAY",
		Report:           fileReport,
		SkuCol:           "E",
		OnHandCol:        "F",
		OnPOCol:          "I",
		OnSOBOCol:        "K",
		YtdSoldIssuedCol: "P",
	}
	holiday := Update{
		Hotsheet:         fileHotsheetNew,
		Sheet:            "HOLIDAY",
		Report:           fileReport,
		SkuCol:           "C",
		OnHandCol:        "D",
		OnPOCol:          "F",
		OnSOBOCol:        "H",
		YtdSoldIssuedCol: "N",
	}

	// Update the hotsheet
	err := everyday.Update("SMD", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update stock: %w", err)
	}
	fmt.Println("Everyday product updated successfully")

	err = holiday.Update("SMD", "holiday")
	if err != nil {
		return fmt.Errorf("failed to update holiday stock: %w", err)
	}
	fmt.Println("Holiday product updated successfully")

	return nil
}
