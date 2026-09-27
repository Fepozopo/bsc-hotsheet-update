package hotsheet

import (
	"fmt"
	"log/slog"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// buildProductLineWorkbook creates one workbook for a product line, writes the standard report
// sheets, Data Insights, Best Sellers, and the optional Monthly History and MTO sheets,
// then saves the result. A nil bestSellersRange uses inventory YTD sales; when
// hasHistory is true, historyRunDate anchors the MTO forecasts.
func buildProductLineWorkbook(productLine string, entries []*inventoryEntry, outputDir, dateStamp string, hasPO, hasHistory bool, bestSellersRange *BestSellersRange, historyRunDate time.Time, logger *slog.Logger) (string, error) {
	f := newProductLineWorkbook()
	defer func() {
		_ = f.Close()
	}()

	if err := writeStandardSheets(f, entries, hasPO); err != nil {
		if logger != nil {
			logger.Error("failed to write standard sheets", "productLine", productLine, "err", err)
		}
		return "", fmt.Errorf("failed to write standard sheets for %s: %w", productLine, err)
	}

	if err := writeDataInsightsSheet(f, entries); err != nil {
		if logger != nil {
			logger.Error("failed to create Data Insights sheet", "productLine", productLine, "err", err)
		}
		return "", fmt.Errorf("failed to create Data Insights sheet for %s: %w", productLine, err)
	}

	if hasHistory {
		if err := writeMonthlyHistorySheet(f, entries); err != nil {
			return "", fmt.Errorf("failed to create Monthly History sheet for %s: %w", productLine, err)
		}
	}

	if err := writeBestSellersSheet(f, entries, bestSellersRange); err != nil {
		return "", fmt.Errorf("failed to create Best Sellers sheet for %s: %w", productLine, err)
	}
	if hasHistory {
		if err := writeMTOSheet(f, entries, historyRunDate); err != nil {
			return "", fmt.Errorf("failed to create MTO sheet for %s: %w", productLine, err)
		}
	}

	outPath, err := saveWorkbook(f, outputDir, productLine, dateStamp)
	if err != nil {
		if logger != nil {
			logger.Error("failed to save hotsheet for product line", "productLine", productLine, "err", err)
		}
		return "", err
	}

	return outPath, nil
}

// newProductLineWorkbook creates the workbook shell used for each product-line export.
func newProductLineWorkbook() *excelize.File {
	f := excelize.NewFile()
	idx, _ := f.NewSheet("Everyday")
	f.SetActiveSheet(idx)
	_, _ = f.NewSheet("Winter")
	_, _ = f.NewSheet("Spring")
	// Delete the default Sheet1 if it still exists so the output matches the existing workbook layout.
	if idxSheet, _ := f.GetSheetIndex("Sheet1"); idxSheet != -1 {
		_ = f.DeleteSheet("Sheet1")
	}
	return f
}

// saveWorkbook builds the final output path, sanitizes the product line name, and writes the
// workbook to disk.
func saveWorkbook(f *excelize.File, outputDir, productLine, dateStr string) (string, error) {
	outDir := outputDir
	if strings.TrimSpace(outDir) == "" {
		outDir = "."
	}
	fileName := fmt.Sprintf("%s_hotsheet_%s.xlsx", sanitizeFileName(productLine), dateStr)
	outPath := filepath.Join(outDir, fileName)
	if err := f.SaveAs(outPath); err != nil {
		return "", fmt.Errorf("failed to save hotsheet %s: %w", outPath, err)
	}
	return outPath, nil
}
