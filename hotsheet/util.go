package hotsheet

import (
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

// Token lists used by mapOccasion
var (
	everTokens = []string{
		"ALL OCCASION",
		"BABY",
		"BAPTISM-COMMUNION",
		"BIRTHDAY",
		"BLANK",
		"CAMP",
		"CANCER",
		"CONGRATULATIONS",
		"ENCOURAGEMENT",
		"FRIENDSHIP",
		"GET WELL",
		"KID BIRTHDAY",
		"LOVE",
		"MENOPAUSE",
		"MISS YOU",
		"NEW HOME",
		"PET SYMPATHY",
		"SORRY",
		"SYMPATHY",
		"TEACHER APPRECIATION",
		"THANK YOU",
		"THINKING OF YOU",
		"WEDDING ANNIVERSARY",
	}
	winterTokens = []string{
		"CHRISTMAS",
		"HALLOWEEN",
		"HANUKKAH",
		"HOLIDAY",
		"THANKSGIVING",
		"VETERAN'S DAY",
		"VETERANS DAY",
	}
	springTokens = []string{
		"EASTER",
		"FATHER'S DAY",
		"FATHERS DAY",
		"GRADUATION",
		"INDEPENDENCE DAY",
		"MOTHER'S DAY",
		"MOTHERS DAY",
		"ST PATRICKS DAY",
		"ST. PATRICK'S DAY",
		"VALENTINE'S DAY",
		"VALENTINES DAY",
	}
)

// entry represents a single inventory item (moved from CreateFromReports)
type entry struct {
	SKU         string
	ProductLine string
	ClassDesc   string
	// RawClassDesc preserves the original inventory category before any display-time prefixing.
	RawClassDesc string
	Status       string
	OnHand       int
	// Per-PO details from the PO report
	PONum1         string
	OnPO1          int
	PONum2         string
	OnPO2          int
	OnPO           int
	OnSO           int
	OnBO           int
	TotalAvailable int
	YTDSold        int
	YTDIssued      int
	SoldPY         int
	IssuedPY       int
	Foil           string
	Occasion       string
	Description    string
	UPC            string
	// Additional fields: royalty and dollar sales (added for new report columns)
	RoyaltyCode   string
	DollarSoldYTD float64
	DollarSoldPY  float64
}

// colToIndex converts Excel column letters to zero-based index (A->0)
func colToIndex(col string) int {
	col = strings.ToUpper(col)
	idx := 0
	for i := 0; i < len(col); i++ {
		idx *= 26
		idx += int(col[i]-'A') + 1
	}
	return idx - 1
}

// parseInt parses numbers permissively (commas, floats fallback, trailing "-" interpreted as negative)
func parseInt(s string) int {
	s = strings.TrimSpace(strings.ReplaceAll(s, ",", ""))
	if s == "" {
		return 0
	}
	if strings.HasSuffix(s, "-") {
		s = "-" + strings.TrimSuffix(s, "-")
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

// parseFloat parses numbers permissively (commas, trailing "-" interpreted as negative)
// Returns a float64, or 0.0 on parse failure.
func parseFloat(s string) float64 {
	s = strings.TrimSpace(strings.ReplaceAll(s, ",", ""))
	if s == "" {
		return 0.0
	}
	if strings.HasSuffix(s, "-") {
		s = "-" + strings.TrimSuffix(s, "-")
	}
	v, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return 0.0
	}
	return v
}

// getCellAt reads a cell from rows by 1-based row number and 0-based column index
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

// isRunDate determines whether a cell looks like a run-date (tries to detect explicit dates/times or a "Run Date" label)
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
	dateRe := regexp.MustCompile(`\b\d{1,2}/\d{1,2}/\d{2,4}\b`)
	if dateRe.MatchString(s) {
		return true
	}
	// match time patterns like 12:34
	timeRe := regexp.MustCompile(`\b\d{1,2}:\d{2}\b`)
	if timeRe.MatchString(s) {
		return true
	}
	return false
}

// getRow returns the row at 1-based index vloc (or nil)
func getRow(rows [][]string, vloc int) []string {
	if vloc-1 >= 0 && vloc-1 < len(rows) {
		return rows[vloc-1]
	}
	return nil
}

// getCell returns the string at index idx in row r (or empty string)
func getCell(r []string, idx int) string {
	if r == nil || idx < 0 || idx >= len(r) {
		return ""
	}
	return r[idx]
}

// assignPO assigns a PO number and quantity into the first available PONum slot on the entry
func assignPO(e *entry, poNum string, qty int) {
	if poNum == "" && qty == 0 {
		return
	}
	if e.PONum1 == "" {
		e.PONum1 = poNum
		e.OnPO1 = qty
		return
	}
	if e.PONum2 == "" {
		e.PONum2 = poNum
		e.OnPO2 = qty
		return
	}
	// fallback: accumulate into OnPO1
	e.OnPO1 += qty
}

// mapOccasion maps raw occasion text to one of: "Everyday", "Winter", "Spring"
func mapOccasion(occ string) string {
	o := strings.ToUpper(strings.TrimSpace(occ))
	if o == "" {
		return "Everyday"
	}
	for _, t := range springTokens {
		if strings.Contains(o, t) {
			return "Spring"
		}
	}
	for _, t := range winterTokens {
		if strings.Contains(o, t) {
			return "Winter"
		}
	}
	for _, t := range everTokens {
		if strings.Contains(o, t) {
			return "Everyday"
		}
	}
	return "Everyday"
}

// simple sanitizer for file names
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
