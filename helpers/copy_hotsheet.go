package helpers

import (
	"fmt"
	"path/filepath"
	"time"

	"github.com/xuri/excelize/v2"
)

// CopyHotsheet creates a copy of the specified hotsheet Excel file,
// saving it with a new name that includes the product name and the
// current date. The function returns the path to the newly created
// file or an error if the operation fails.
//
// Parameters:
//   - product: A string representing the product name to be included
//     in the new file name.
//   - hotsheet: A string representing the path to the existing hotsheet
//     Excel file to be copied.
//
// Returns:
//   - A string representing the path to the newly created hotsheet file.
//   - An error if the file cannot be opened or saved.
func CopyHotsheet(product, hotsheet string) (string, error) {
	// Get the current date for hotsheet name
	currentDateHotsheet := time.Now().Format("2006-01-02")

	// Open the hotsheet workbook
	wbHotsheet, err := excelize.OpenFile(hotsheet)
	if err != nil {
		return "", fmt.Errorf("failed to open hotsheet file %s: %w", hotsheet, err)
	}
	defer wbHotsheet.Close()

	// Create a new file path
	fileName := fmt.Sprintf("%s_hotsheet_%v.xlsx", product, currentDateHotsheet)
	baseDir := filepath.Dir(hotsheet)
	filePath := filepath.Join(baseDir, fileName)

	// Copy the hotsheet with a different name
	err = wbHotsheet.SaveAs(filePath)
	if err != nil {
		return "", fmt.Errorf("failed to save hotsheet file %s: %w", filePath, err)
	}

	return filePath, nil
}
