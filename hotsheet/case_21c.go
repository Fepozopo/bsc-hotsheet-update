package hotsheet

import (
	"fmt"
)

// Case21C is a helper function for updating the 21C hotsheet. It creates an Update struct
// and calls its Update method.
//
// This function is called by the main function when the user selects '21C' as the product
// line.
func Case21C(fileHotsheetNew, fileReport string) error {
	everyday := Update{
		Hotsheet:     fileHotsheetNew,
		Sheet:        "EVERYDAY",
		Report:       fileReport,
		SkuCol:       "C",
		OnHandCol:    "D",
		OnPOCol:      "E",
		OnSOBOCol:    "G",
		YtdSoldCol:   "M",
		YtdIssuedCol: "N",
	}

	// Update the hotsheet
	err := everyday.Update("21c", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update stock: %w", err)
	}
	fmt.Println("Everyday product updated successfully")

	return nil
}
