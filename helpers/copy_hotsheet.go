package helpers

import (
	"fmt"
	"path/filepath"
	"time"

	"github.com/xuri/excelize/v2"
)

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
	fileName := fmt.Sprintf("%s_Hotsheet_%v.xlsx", product, currentDateHotsheet)
	baseDir := filepath.Dir(hotsheet)
	filePath := filepath.Join(baseDir, fileName)

	// Copy the hotsheet with a different name
	err = wbHotsheet.SaveAs(filePath)
	if err != nil {
		return "", fmt.Errorf("failed to save hotsheet file %s: %w", filePath, err)
	}

	return filePath, nil
}
