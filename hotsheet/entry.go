package hotsheet

// inventoryEntry represents a single inventory item.
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
	// Additional fields: royalty and dollar sales (added for new report columns)
	RoyaltyCode   string
	DollarSoldYTD float64
	DollarSoldPY  float64
}
