// Workbook scheduling keeps Excelize files independent and progress callbacks serial.
package hotsheet

import (
	"fmt"
	"sort"
	"strings"
)

// maxConcurrentWorkbooks limits overlapping workbook object graphs and XML buffers.
const maxConcurrentWorkbooks = 2

// workbookResult carries a completed writer's path or error and its input position.
type workbookResult struct {
	path  string
	err   error
	index int
}

// writeProductLineWorkbooks calls build for each product line, with at most two
// calls active. build must own each line's mutable data and Excelize file. report
// runs only on this function's calling goroutine, never on a writer goroutine.
// Colliding sanitized filenames are rejected before build runs. Successful paths
// retain productLines' order. After the first observed build error, no new work
// starts; active calls finish before successful paths and that error are returned.
func writeProductLineWorkbooks(productLines []string, build func(string) (string, error), report ProgressCallback) ([]string, error) {
	// Case folding also protects case-insensitive destination filesystems.
	// All outputs share the same directory, date suffix, and file extension.
	names := make(map[string]string, len(productLines))
	for _, line := range productLines {
		name := strings.ToLower(sanitizeFileName(line))
		if previous, found := names[name]; found {
			return nil, fmt.Errorf("%w: %q and %q", ErrProductLineFilenameCollision, previous, line)
		}
		names[name] = line
	}

	paths := make([]string, len(productLines))
	completed := make([]int, 0, len(productLines))
	results := make(chan workbookResult, maxConcurrentWorkbooks)
	next, active := 0, 0
	// Only the coordinator starts work and reports progress, preserving the
	// callback contract even when writers finish in a different order.
	startNext := func() {
		index := next
		line := productLines[index]
		reportGenerationProgress(report, workbookProgress(len(completed), len(productLines)), fmt.Sprintf("Writing %s hotsheet...", line))
		next++
		active++
		go func() {
			path, err := build(line)
			results <- workbookResult{path: path, err: err, index: index}
		}()
	}
	for active < maxConcurrentWorkbooks && next < len(productLines) {
		startNext()
	}

	var firstErr error
	for active > 0 {
		result := <-results
		active--
		if result.err != nil {
			if firstErr == nil {
				firstErr = result.err
			}
		} else {
			paths[result.index] = result.path
			completed = append(completed, result.index)
			reportGenerationProgress(report, workbookProgress(len(completed), len(productLines)), fmt.Sprintf("Created %d of %d hotsheets.", len(completed), len(productLines)))
		}
		if firstErr == nil && next < len(productLines) {
			startNext()
		}
	}

	// Drain every active writer before returning: Generate can now close its
	// shared logger, and callers can safely inspect all reported output files.
	sort.Ints(completed)
	outputs := make([]string, 0, len(completed))
	for _, index := range completed {
		outputs = append(outputs, paths[index])
	}
	return outputs, firstErr
}
