package hotsheet

// inventoryEntry represents a single inventory item and holds the source-report and derived values
// used to build its hotsheet rows, including optional BSC monthly sales and issue history.
type inventoryEntry struct {
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
	RoyaltyCode    string
	DollarSoldYTD  float64
	DollarSoldPY   float64
	CardSize       string
	Inactive       string
	// SalesRecords contain BSC warehouse year/metric rows from optional sales and
	// issue history; issued values are clamped to zero per month at import.
	SalesRecords []salesRecord
}

// salesRecord holds a BSC sales or issued metric for one calendar year, with
// periods 1–12 mapped to January–December. Sales values retain the source scale
// (including 0–100 for gross profit percentages); issued values are nonnegative.
type salesRecord struct {
	Year    int
	Metric  string
	Periods [12]float64
}
