package main

import (
	"fmt"

	"github.com/sqweek/dialog"
)

func handlerGetFiles(product string) (string, string, string, error) {
	fileHotsheet, err := dialog.File().Title("Select the HOTSHEET...").Filter("Excel Files", "*.xlsx").Load()
	if err != nil {
		return "", "", "", fmt.Errorf("failed to open hotsheet file: %w", err)
	}
	fileStockReport, err := dialog.File().Title("Select the Stock Report...").Filter("Excel Files", "*.xlsx").Load()
	if err != nil {
		return "", "", "", fmt.Errorf("failed to open stock report file: %w", err)
	}
	fileSalesReport, err := dialog.File().Title("Select the Sales Report...").Filter("Excel Files", "*.xlsx").Load()
	if err != nil {
		return "", "", "", fmt.Errorf("failed to open sales report file: %w", err)
	}
	fileHotsheetNew, err := handlerCopyHotsheet(product, fileHotsheet)
	if err != nil {
		return "", "", "", fmt.Errorf("failed to copy hotsheet file: %w", err)
	}

	return fileHotsheetNew, fileStockReport, fileSalesReport, nil
}
