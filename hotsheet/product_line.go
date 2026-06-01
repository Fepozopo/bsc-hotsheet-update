package hotsheet

import (
	"log/slog"
	"sort"
	"strings"
)

// groupEntriesByProductLine iterates the inventory map and groups entries by ProductLine while
// preserving the current behavior of skipping blank product-line entries.
func groupEntriesByProductLine(inventoryBySKU map[string]*inventoryEntry, logger *slog.Logger) map[string][]*inventoryEntry {
	productGroups := make(map[string][]*inventoryEntry)
	for _, item := range inventoryBySKU {
		if item == nil {
			continue
		}
		pl := strings.TrimSpace(item.ProductLine)
		if pl == "" {
			if logger != nil {
				// Skip entries without a ProductLine so the workbook does not create UNKNOWN files.
				logger.Info("Skipping SKU with empty ProductLine (likely PO-only entry)", "SKU", item.SKU)
			}
			continue
		}
		productGroups[pl] = append(productGroups[pl], item)
	}
	return productGroups
}

// sortEntriesForProductLine sorts a product-line slice into a stable SKU order before workbook
// generation so the tab contents remain consistent.
func sortEntriesForProductLine(entries []*inventoryEntry) {
	sort.SliceStable(entries, func(i, j int) bool { return entries[i].SKU < entries[j].SKU })
}
