// Optional reports are parsed independently before their data is merged into inventory.
package hotsheet

import (
	"fmt"
	"log/slog"
	"strings"
	"time"
)

// maxConcurrentImports bounds overlapping Excelize readers and their XML buffers.
// Workbook generation starts only after every reader has finished.
const maxConcurrentImports = 2

const (
	salesImport = iota
	issueImport
	poImport
	optionalImportCount
)

// reportReader parses reports into private buffers. Implementations may be called
// concurrently and must not modify inventory or invoke progress callbacks.
type reportReader interface {
	readHistory(path string, issue bool) (monthlyHistoryData, error)
	readPO(path string) (poImportData, error)
}

// excelReportReader opens an independent Excelize workbook for each report.
// The shared inventory map is used only to select matching, non-nil entries.
type excelReportReader struct {
	inventoryBySKU map[string]*inventoryEntry
	logger         *slog.Logger
}

// readHistory parses matching BSC records and the report date without attaching
// them to inventory. issue selects issued quantities instead of sales metrics.
func (reader excelReportReader) readHistory(path string, issue bool) (monthlyHistoryData, error) {
	return readMonthlyHistory(path, reader.inventoryBySKU, reader.logger, issue)
}

// readPO parses matching PO numbers and quantities without changing inventory.
func (reader excelReportReader) readPO(path string) (poImportData, error) {
	return readPOData(path, reader.inventoryBySKU, reader.logger)
}

// reportReadTask describes one optional reader call and its user-facing label.
type reportReadTask struct {
	path string
	name string
	read func() error
}

// reportReadResult identifies a finished task and its parse error, if any.
type reportReadResult struct {
	err   error
	index int
}

// mergeOptionalReports parses supplied reports with two slots, prioritizing sales
// and issue history. PO takes the first freed slot. reader returns private buffers;
// report is invoked synchronously on this function's calling goroutine.
// All readers finish before data is merged in PO, sales, issue order or an error
// is returned. PO errors are logged and ignored; sales errors take precedence over
// issue errors regardless of completion order. The returned date is from sales.
func mergeOptionalReports(poPath, salesPath, issuePath string, reader reportReader, logger *slog.Logger, report ProgressCallback) (time.Time, error) {
	var sales, issue monthlyHistoryData
	var po poImportData
	tasks := [optionalImportCount]reportReadTask{
		salesImport: {path: salesPath, name: "monthly sales history", read: func() error {
			var err error
			sales, err = reader.readHistory(salesPath, false)
			return err
		}},
		issueImport: {path: issuePath, name: "monthly issue history", read: func() error {
			var err error
			issue, err = reader.readHistory(issuePath, true)
			return err
		}},
		poImport: {path: poPath, name: "PO report", read: func() error {
			var err error
			po, err = reader.readPO(poPath)
			return err
		}},
	}
	pending := make([]int, 0, len(tasks))
	for index, task := range tasks {
		if strings.TrimSpace(task.path) != "" {
			pending = append(pending, index)
		}
	}
	if len(pending) == 0 {
		return time.Time{}, nil
	}

	// Progress counts finished reads, not start order: PO can start after a
	// history report finishes without moving the percentage backwards.
	const importStart, importEnd = 30, 44
	completed, next, active := 0, 0, 0
	results := make(chan reportReadResult, maxConcurrentImports)
	startNext := func() {
		index := pending[next]
		percent := importStart + (importEnd-importStart)*completed/len(pending)
		reportGenerationProgress(report, percent, fmt.Sprintf("Loading %s...", tasks[index].name))
		next++
		active++
		go func() {
			err := tasks[index].read()
			results <- reportReadResult{err: err, index: index}
		}()
	}
	for active < maxConcurrentImports && next < len(pending) {
		startNext()
	}
	var readErrors [optionalImportCount]error
	for active > 0 {
		result := <-results
		active--
		completed++
		readErrors[result.index] = result.err
		message := fmt.Sprintf("Loaded %s.", tasks[result.index].name)
		if result.err != nil {
			message = fmt.Sprintf("Failed to load %s.", tasks[result.index].name)
		}
		percent := importStart + (importEnd-importStart)*completed/len(pending)
		reportGenerationProgress(report, percent, message)
		if next < len(pending) {
			startNext()
		}
	}

	// Receiving every result joins the readers before inventory mutation and
	// before Generate can close the logger, including on history failures.
	po.merge()
	if err := readErrors[poImport]; err != nil && logger != nil {
		logger.Error("failed to merge PO report", "err", err)
	}
	sales.merge()
	if err := readErrors[salesImport]; err != nil {
		return time.Time{}, err
	}
	issue.merge()
	if err := readErrors[issueImport]; err != nil {
		return time.Time{}, err
	}
	return sales.runDate, nil
}
