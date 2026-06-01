package hotsheet

import (
	"fmt"
	"log/slog"
	"time"

	"github.com/Fepozopo/bsc-hotsheet-update/helpers"
)

// Generate orchestrates the hotsheet report pipeline. It loads the source data,
// merges optional PO information, groups entries by product line, and delegates
// workbook generation to focused helpers.
func Generate(inventoryPath, poPath, outputDir string) ([]string, error) {
	logger, logCloser, err := newReportLogger()
	if err != nil {
		return nil, err
	}
	defer func() {
		_ = logCloser.Close()
	}()

	logger.Info("hotsheet generation started", "inventoryPath", inventoryPath, "poPath", poPath, "outputDir", outputDir)

	inventoryBySKU, err := loadInventoryEntries(inventoryPath, logger)
	if err != nil {
		return nil, err
	}

	hasPO := poPath != ""
	if hasPO {
		if err := mergePOData(poPath, inventoryBySKU, logger); err != nil {
			logger.Error("failed to merge PO report", "err", err)
		}
	}

	entriesByProductLine := groupEntriesByProductLine(inventoryBySKU, logger)
	outputs := make([]string, 0, len(entriesByProductLine))
	dateStamp := currentDateStamp()

	for productLine, entries := range entriesByProductLine {
		sortEntriesForProductLine(entries)

		outPath, err := buildProductLineWorkbook(productLine, entries, outputDir, dateStamp, hasPO, logger)
		if err != nil {
			return outputs, err
		}
		outputs = append(outputs, outPath)
	}

	logger.Info("hotsheet generation completed", "filesCreated", len(outputs), "outputDir", outputDir)
	return outputs, nil
}

// newReportLogger constructs the logger used by Generate so the orchestration layer
// stays focused on the report pipeline itself.
func newReportLogger() (*slog.Logger, interface{ Close() error }, error) {
	logger, logCloser, err := helpers.CreateSlogLogger("create", "DEBUG")
	if err != nil {
		return nil, nil, fmt.Errorf("failed to create logger: %w", err)
	}
	return logger, logCloser, nil
}

// currentDateStamp returns the date stamp used in output filenames.
func currentDateStamp() string {
	return time.Now().Format("20060102")
}
