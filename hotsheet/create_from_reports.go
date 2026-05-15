package hotsheet

import (
	"fmt"
	"log/slog"
	"time"

	"github.com/Fepozopo/bsc-hotsheet-update/helpers"
)

// CreateFromReports exists as the orchestration layer for the report pipeline. It loads the
// source data, merges optional PO information, groups entries by product line, and delegates
// workbook generation to focused helpers.
func CreateFromReports(inventoryPath, poPath, outputDir string) ([]string, error) {
	logger, logCloser, err := createReportLogger()
	if err != nil {
		return nil, err
	}
	defer func() {
		_ = logCloser.Close()
	}()

	logger.Info("CreateFromReports started", "inventoryPath", inventoryPath, "poPath", poPath, "outputDir", outputDir)

	invMap, err := loadInventoryEntries(inventoryPath, logger)
	if err != nil {
		return nil, err
	}

	hasPO := poPath != ""
	if hasPO {
		if err := mergePOData(poPath, invMap, logger); err != nil {
			logger.Error("failed to merge PO report", "err", err)
		}
	}

	productGroups := groupEntriesByProductLine(invMap, logger)
	outputs := make([]string, 0, len(productGroups))
	dateStr := currentDateString()

	for productLine, entries := range productGroups {
		sortEntriesForProductLine(entries)

		outPath, err := buildProductLineWorkbook(productLine, entries, outputDir, dateStr, hasPO, logger)
		if err != nil {
			return outputs, err
		}
		outputs = append(outputs, outPath)
	}

	logger.Info("CreateFromReports completed", "filesCreated", len(outputs), "outputDir", outputDir)
	return outputs, nil
}

// createReportLogger constructs the logger used by CreateFromReports so the orchestration layer
// stays focused on the report pipeline itself.
func createReportLogger() (*slog.Logger, interface{ Close() error }, error) {
	logger, logCloser, err := helpers.CreateSlogLogger("create", "all", "", "DEBUG")
	if err != nil {
		return nil, nil, fmt.Errorf("failed to create logger: %w", err)
	}
	return logger, logCloser, nil
}

// currentDateString returns the date stamp used in output filenames.
func currentDateString() string {
	return time.Now().Format("20060102")
}
