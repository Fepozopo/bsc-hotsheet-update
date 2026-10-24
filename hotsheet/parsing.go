package hotsheet

import (
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

// Compile constant footer patterns once: inventory parsing checks every SKU.
var (
	runDatePattern = regexp.MustCompile(`\b\d{1,2}/\d{1,2}/\d{2,4}\b`)
	runTimePattern = regexp.MustCompile(`\b\d{1,2}:\d{2}\b`)
)

// colToIndex converts Excel column letters to zero-based index (A->0).
func colToIndex(col string) int {
	col = strings.ToUpper(col)
	idx := 0
	for i := 0; i < len(col); i++ {
		idx *= 26
		idx += int(col[i]-'A') + 1
	}
	return idx - 1
}

// parseInt parses s permissively, allowing commas, a float fallback, and a trailing
// "-" to indicate a negative number. It returns 0 on parse failure.
func parseInt(s string) int {
	s = strings.TrimSpace(strings.ReplaceAll(s, ",", ""))
	if s == "" {
		return 0
	}
	if prefix, found := strings.CutSuffix(s, "-"); found {
		s = "-" + prefix
	}
	v, err := strconv.Atoi(s)
	if err != nil {
		f, err2 := strconv.ParseFloat(s, 64)
		if err2 != nil {
			return 0
		}
		return int(f)
	}
	return v
}

// parseFloat parses s permissively, allowing commas and a trailing "-" to indicate
// a negative number. It returns 0.0 on parse failure.
func parseFloat(s string) float64 {
	s = strings.TrimSpace(strings.ReplaceAll(s, ",", ""))
	if s == "" {
		return 0.0
	}
	if prefix, found := strings.CutSuffix(s, "-"); found {
		s = "-" + prefix
	}
	v, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return 0.0
	}
	return v
}

// getCellAt reads a cell from rows by 1-based row number and 0-based column index.
func getCellAt(rows [][]string, rowNum int, colIdx int) string {
	if rowNum-1 < 0 || rowNum-1 >= len(rows) {
		return ""
	}
	row := rows[rowNum-1]
	if colIdx < 0 || colIdx >= len(row) {
		return ""
	}
	return strings.TrimSpace(row[colIdx])
}

// isRunDate reports whether s contains a run-date label, a slash-separated date,
// or a time. Shared compiled patterns preserve footer detection without repeated
// compilation for each inventory SKU.
func isRunDate(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return false
	}
	upper := strings.ToUpper(s)
	// explicit label like "Run Date" is a strong indicator
	if strings.Contains(upper, "RUN") && strings.Contains(upper, "DATE") {
		return true
	}
	// match common date formats like 04/09/2026 or 4/9/26
	if runDatePattern.MatchString(s) {
		return true
	}
	// match time patterns like 12:34
	return runTimePattern.MatchString(s)
}

// getRow returns the row at 1-based index vloc (or nil).
func getRow(rows [][]string, vloc int) []string {
	if vloc-1 >= 0 && vloc-1 < len(rows) {
		return rows[vloc-1]
	}
	return nil
}

// getCell returns the string at index idx in row r (or empty string).
func getCell(r []string, idx int) string {
	if r == nil || idx < 0 || idx >= len(r) {
		return ""
	}
	return r[idx]
}

// simple sanitizer for file names. It strips path separators and colons so the generated
// output names stay safe on the local filesystem.
func sanitizeFileName(s string) string {
	s = strings.TrimSpace(s)
	if s == "" {
		return "unknown"
	}
	// replace path separators and colons
	s = strings.ReplaceAll(s, string(filepath.Separator), "_")
	s = strings.ReplaceAll(s, ":", "")
	return s
}
