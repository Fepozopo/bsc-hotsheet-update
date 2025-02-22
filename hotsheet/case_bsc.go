package hotsheet

import (
	"fmt"
)

// CaseBSC is a helper function for updating the BSC hotsheet. It creates an Update struct
// and calls its Update method.
//
// This function is called by the main function when the user selects 'BSC' as the product
// line.
func CaseBSC(fileHotsheetNew, fileReport string) error {
	everyday := Update{
		Hotsheet:     fileHotsheetNew,
		Sheet:        "Everyday",
		Report:       fileReport,
		SkuCol:       "D",
		OnHandCol:    "E",
		OnPOCol:      "F",
		OnSOBOCol:    "H",
		YtdSoldCol:   "K",
		YtdIssuedCol: "L",
	}
	winter := Update{
		Hotsheet:     fileHotsheetNew,
		Sheet:        "Winter Holiday",
		Report:       fileReport,
		SkuCol:       "E",
		OnHandCol:    "F",
		OnPOCol:      "I",
		OnSOBOCol:    "G",
		YtdSoldCol:   "L",
		YtdIssuedCol: "M",
	}
	spring := Update{
		Hotsheet:     fileHotsheetNew,
		Sheet:        "Spring Holiday",
		Report:       fileReport,
		SkuCol:       "D",
		OnHandCol:    "E",
		OnPOCol:      "H",
		OnSOBOCol:    "J",
		YtdSoldCol:   "L",
		YtdIssuedCol: "M",
	}
	a2Notecards := Update{
		Hotsheet:     fileHotsheetNew,
		Sheet:        "A2 Notecards",
		Report:       fileReport,
		SkuCol:       "D",
		OnHandCol:    "F",
		OnPOCol:      "G",
		OnSOBOCol:    "I",
		YtdSoldCol:   "L",
		YtdIssuedCol: "M",
	}

	// Update the hotsheet
	err := everyday.Update("BSC", "everyday")
	if err != nil {
		return fmt.Errorf("failed to update stock: %w", err)
	}
	fmt.Println("Everyday product updated successfully")

	err = winter.Update("BSC", "winter")
	if err != nil {
		return fmt.Errorf("failed to update holiday stock: %w", err)
	}
	fmt.Println("Holiday product updated successfully")

	err = spring.Update("BSC", "spring")
	if err != nil {
		return fmt.Errorf("failed to update spring stock: %w", err)
	}
	fmt.Println("Spring product updated successfully")

	err = a2Notecards.Update("BSC", "a2notecards")
	if err != nil {
		return fmt.Errorf("failed to update A2 Notecards stock: %w", err)
	}
	fmt.Println("A2 Notecard product updated successfully")

	return nil
}
