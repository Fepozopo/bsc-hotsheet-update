package hotsheet

import (
	"fmt"
	"log/slog"
	"time"

	"github.com/Fepozopo/bsc-hotsheet-update/helpers"
)

// Progress describes a user-visible generation milestone.
//
// Progress is intentionally coarse-grained: generation does not currently know
// the exact amount of work required for every Excel row and cell ahead of time,
// so the reported percentage is based on major pipeline phases plus completed
// product-line workbooks. That gives the GUI a determinate bar that reflects
// real completed work without coupling the UI to every low-level writer loop.
type Progress struct {
	// Percent is the overall completion percentage, clamped by the generator to
	// the inclusive range 0..100 before it is reported to callers.
	Percent int

	// Message is a short status string suitable for displaying next to a progress
	// bar, such as "Loading inventory report..." or "Created 2 of 5 hotsheets.".
	Message string
}

// ProgressCallback receives generation progress updates.
//
// Callbacks are invoked synchronously from the generation goroutine. Callers
// that update UI state should marshal the Progress value onto their UI thread
// instead of mutating UI-owned state directly.
type ProgressCallback func(Progress)

// Generate orchestrates the hotsheet report pipeline.
//
// It loads the source inventory data, merges optional PO information, groups
// entries by product line, and writes one workbook per product line. If report is
// non-nil, Generate reports determinate progress at major pipeline milestones
// and after each product-line workbook is written. Passing nil disables progress
// reporting.
func Generate(inventoryPath, poPath, outputDir string, report ProgressCallback) ([]string, error) {
	reportGenerationProgress(report, 0, "Starting generation...")

	logger, logCloser, err := newReportLogger()
	if err != nil {
		return nil, err
	}
	defer func() {
		_ = logCloser.Close()
	}()

	logger.Info("hotsheet generation started", "inventoryPath", inventoryPath, "poPath", poPath, "outputDir", outputDir)
	reportGenerationProgress(report, 5, "Loading inventory report...")

	inventoryBySKU, err := loadInventoryEntries(inventoryPath, logger)
	if err != nil {
		return nil, err
	}
	reportGenerationProgress(report, 30, "Inventory report loaded.")

	hasPO := poPath != ""
	if hasPO {
		reportGenerationProgress(report, 35, "Merging PO report...")
		if err := mergePOData(poPath, inventoryBySKU, logger); err != nil {
			logger.Error("failed to merge PO report", "err", err)
		}
	}
	reportGenerationProgress(report, 45, "Grouping product lines...")

	entriesByProductLine := groupEntriesByProductLine(inventoryBySKU, logger)
	outputs := make([]string, 0, len(entriesByProductLine))
	dateStamp := currentDateStamp()
	totalProductLines := len(entriesByProductLine)
	if totalProductLines == 0 {
		reportGenerationProgress(report, 100, "Generation complete.")
		logger.Info("hotsheet generation completed", "filesCreated", len(outputs), "outputDir", outputDir)
		return outputs, nil
	}

	for productLine, entries := range entriesByProductLine {
		reportGenerationProgress(report, workbookProgress(len(outputs), totalProductLines), fmt.Sprintf("Writing %s hotsheet...", productLine))
		sortEntriesForProductLine(entries)

		outPath, err := buildProductLineWorkbook(productLine, entries, outputDir, dateStamp, hasPO, logger)
		if err != nil {
			return outputs, err
		}
		outputs = append(outputs, outPath)
		reportGenerationProgress(report, workbookProgress(len(outputs), totalProductLines), fmt.Sprintf("Created %d of %d hotsheets.", len(outputs), totalProductLines))
	}

	reportGenerationProgress(report, 100, "Generation complete.")
	logger.Info("hotsheet generation completed", "filesCreated", len(outputs), "outputDir", outputDir)
	return outputs, nil
}

// reportGenerationProgress normalizes and emits a Progress update.
//
// Keeping the nil check and percent clamping in one helper makes each generation
// milestone easy to read and prevents future call sites from accidentally
// sending invalid percentages to the GUI.
func reportGenerationProgress(report ProgressCallback, percent int, message string) {
	if report == nil {
		return
	}
	if percent < 0 {
		percent = 0
	}
	if percent > 100 {
		percent = 100
	}
	report(Progress{Percent: percent, Message: message})
}

// workbookProgress maps completed product-line workbooks into the percentage
// range reserved for workbook generation.
//
// Earlier stages reserve 0..49 for setup and input processing. Workbook writing
// uses 50..95 so the bar advances proportionally for each output file while
// leaving 100 for the explicit "Generation complete" milestone after all final
// bookkeeping and logging has finished.
func workbookProgress(completed, total int) int {
	if total <= 0 {
		return 100
	}
	const (
		workbookStart = 50
		workbookEnd   = 95
	)
	return workbookStart + (workbookEnd-workbookStart)*completed/total
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
