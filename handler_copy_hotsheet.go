package main

import (
	"fmt"
	"path/filepath"
	"time"

	"github.com/xuri/excelize/v2"
)

func handlerCopyHotsheet(product, hotsheet string) (string, error) {
	// Open the hotsheet workbook
	wbHotsheet, err := excelize.OpenFile(hotsheet)
	if err != nil {
		return "", err
	}
	defer wbHotsheet.Close()

	// Get the current date
	currentDate := time.Now().Format("2006-01-02")

	// Create a new file path
	fileName := fmt.Sprintf("%s_Hotsheet_%v.xlsx", product, currentDate)
	baseDir := filepath.Dir(hotsheet)
	filePath := filepath.Join(baseDir, fileName)

	// Copy the hotsheet with a different name
	err = wbHotsheet.SaveAs(filePath)
	if err != nil {
		return "", err
	}

	return filePath, nil
}
