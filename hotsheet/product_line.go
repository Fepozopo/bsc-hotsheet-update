package hotsheet

import (
	"log/slog"
	"sort"
	"strings"
)

// groupEntriesByProductLine iterates the inventory map and groups entries by ProductLine while
// preserving the current behavior of skipping blank product-line entries.
func groupEntriesByProductLine(invMap map[string]*entry, logger *slog.Logger) map[string][]*entry {
	productGroups := make(map[string][]*entry)
	for _, e := range invMap {
		if e == nil {
			continue
		}
		pl := strings.TrimSpace(e.ProductLine)
		if pl == "" {
			if logger != nil {
				// Skip entries without a ProductLine so the workbook does not create UNKNOWN files.
				logger.Info("Skipping SKU with empty ProductLine (likely PO-only entry)", "SKU", e.SKU)
			}
			continue
		}
		productGroups[pl] = append(productGroups[pl], e)
	}
	return productGroups
}

// sortEntriesForProductLine sorts a product-line slice into a stable SKU order before workbook
// generation so the tab contents remain consistent.
func sortEntriesForProductLine(entries []*entry) {
	sort.SliceStable(entries, func(i, j int) bool { return entries[i].SKU < entries[j].SKU })
}
