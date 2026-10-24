package hotsheet

import (
	"fmt"
	"log/slog"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// buildProductLineWorkbook creates and saves one product-line workbook, starting
// with Dashboard, then All Products, YTD Stock Priority, Data Insights, Best Sellers,
// and optional Monthly History and MTO. A nil bestSellersRange uses inventory YTD
// shipments. With history, historyRunDate anchors MTO forecasts, Data Insights
// projections, and Dashboard comparisons to completed months. MTO helper data
// and Monthly History are streamed after operations that would reload their XML
// into memory, with tab positions reserved earlier. It returns the saved path
// or an error.
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

	if err := writeDataInsightsSheet(f, entries, hasHistory, historyRunDate); err != nil {
		if logger != nil {
			logger.Error("failed to create Data Insights sheet", "productLine", productLine, "err", err)
		}
		return "", fmt.Errorf("failed to create Data Insights sheet for %s: %w", productLine, err)
	}

	if err := writeBestSellersSheet(f, entries, bestSellersRange); err != nil {
		return "", fmt.Errorf("failed to create Best Sellers sheet for %s: %w", productLine, err)
	}
	var mtoRows []mtoForecastRow
	if hasHistory {
		// Reserve the tab position now; the history writer will reuse this empty
		// sheet after the helper-sheet visibility changes are complete.
		if _, err := f.NewSheet(monthlyHistorySheetName); err != nil {
			return "", fmt.Errorf("failed to create Monthly History sheet for %s: %w", productLine, err)
		}
		mtoRows = buildMTORows(entries, historyRunDate)
		if err := writeMTOSheet(f, mtoRows, historyRunDate); err != nil {
			return "", fmt.Errorf("failed to create MTO sheet for %s: %w", productLine, err)
		}
	}

	if err := writeDashboardSheet(f, entries, hasHistory, historyRunDate); err != nil {
		return "", fmt.Errorf("failed to create Dashboard sheet for %s: %w", productLine, err)
	}

	// Excelize's SetSheetVisible reads every worksheet. Stream both large tabs
	// only after MTO and Dashboard hide their helpers to avoid materializing XML.
	if hasHistory {
		if err := writeMTOForecastDataSheet(f, mtoRows, historyRunDate); err != nil {
			return "", fmt.Errorf("failed to write MTO forecast data for %s: %w", productLine, err)
		}
		if err := writeMonthlyHistorySheet(f, entries); err != nil {
			return "", fmt.Errorf("failed to create Monthly History sheet for %s: %w", productLine, err)
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

// newProductLineWorkbook returns a workbook with Dashboard first and active,
// followed by All Products and YTD Stock Priority. Later writers append report
// tabs and hidden chart/forecast data without changing the initial tab order.
func newProductLineWorkbook() *excelize.File {
	f := excelize.NewFile()
	_ = f.SetSheetName("Sheet1", dashboardSheetName)
	_, _ = f.NewSheet(allProductsSheetName)
	_, _ = f.NewSheet(ytdStockPrioritySheetName)
	f.SetActiveSheet(0)
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
