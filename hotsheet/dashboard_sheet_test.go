package hotsheet

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"path/filepath"
	"slices"
	"strings"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// dashboardSeriesXML reads name/category/value references from saved chart series so
// tests evaluate the actual ranges plotted by Excel, not parallel test formulas.
type dashboardSeriesXML struct {
	Name       string `xml:"tx>strRef>f"`
	Categories string `xml:"cat>strRef>f"`
	Values     string `xml:"val>numRef>f"`
}

// dashboardChartXML reads the chart types and series that users receive in XLSX.
type dashboardChartXML struct {
	Chart struct {
		PlotArea struct {
			Donuts []struct {
				Series []dashboardSeriesXML `xml:"ser"`
			} `xml:"doughnutChart"`
			Bars []struct {
				Series []dashboardSeriesXML `xml:"ser"`
			} `xml:"barChart"`
		} `xml:"plotArea"`
	} `xml:"chart"`
}

// TestDashboardSavedWorkbook verifies the initially active Dashboard, optional
// monthly chart with year legends, independent validated selectors, and class-sensitive chart data
// after saving/reopening. Comparison labels distinguish completed months from
// inventory YTD. The reporting date is fixed rather than using a clock.
func TestDashboardSavedWorkbook(t *testing.T) {
	for _, tc := range []struct {
		name                         string
		hasHistory                   bool
		month                        time.Month
		wantSales, wantFilteredSales []string
	}{
		{"inventory only", false, time.September, []string{"400", "1300"}, []string{"100", "500"}},
		{"September completed months", true, time.September, []string{"110", "77"}, []string{"30", "21"}},
		{"January no completed months", true, time.January, []string{"0", "0"}, []string{"0", "0"}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			entries := []*inventoryEntry{
				{SKU: "A", Description: "Alpha", ClassDesc: "Cards", YTDSold: 8, YTDIssued: 2, DollarSoldYTD: 300, DollarSoldPY: 800, SalesRecords: []salesRecord{
					{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{10, 0, 0, 0, 0, 0, 0, 70, 900}},
					{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{7, 0, 0, 0, 0, 0, 0, 49, 800}},
				}},
				{SKU: "B", Description: "Beta", ClassDesc: "CardsX", YTDSold: 20, YTDIssued: -4, DollarSoldYTD: 100, DollarSoldPY: 500, SalesRecords: []salesRecord{
					{Year: 2026, Metric: "Dollars Sold", Periods: [12]float64{20, 0, 0, 0, 0, 0, 0, 10, 100}},
					{Year: 2025, Metric: "Dollars Sold", Periods: [12]float64{14, 0, 0, 0, 0, 0, 0, 7, 200}},
				}},
			}
			path, err := buildProductLineWorkbook("BAS", entries, t.TempDir(), "20260925", false, tc.hasHistory, nil, time.Date(2026, tc.month, 25, 0, 0, 0, 0, time.UTC), nil)
			if err != nil {
				t.Fatal(err)
			}
			f, err := excelize.OpenFile(path)
			if err != nil {
				t.Fatal(err)
			}
			t.Cleanup(func() { _ = f.Close() })
			if sheets := f.GetSheetList(); sheets[0] != "Dashboard" || f.GetActiveSheetIndex() != 0 {
				t.Fatalf("expected Dashboard first and active, got %v active=%d", sheets, f.GetActiveSheetIndex())
			}
			if visible, err := f.GetSheetVisible("Dashboard Data"); err != nil || visible {
				t.Fatalf("expected hidden chart data, got visible=%v error=%v", visible, err)
			}
			selectors := []string{"B6", "L6", "B50"}
			if tc.hasHistory {
				selectors = append(selectors, "L50")
			}
			validations, err := f.GetDataValidations("Dashboard")
			if err != nil || len(validations) != len(selectors) {
				t.Fatalf("expected %d chart selectors, got %v error=%v", len(selectors), validations, err)
			}
			for _, selector := range selectors {
				value, err := f.GetCellValue("Dashboard", selector)
				if err != nil || value != "All" {
					t.Errorf("selector %s expected All, got %q error=%v", selector, value, err)
				}
				index := slices.IndexFunc(validations, func(v *excelize.DataValidation) bool { return v.Sqref == selector })
				if index == -1 || validations[index].Type != "list" || validations[index].Formula1 != "DashboardClasses" {
					t.Fatalf("selector %s expected class-list dropdown, got %v", selector, validations)
				}
			}
			classesFound := false
			for _, name := range f.GetDefinedName() {
				if name.Name == "DashboardClasses" {
					assertDashboardRange(t, f, name.RefersTo, []string{"All", "Cards", "CardsX"})
					classesFound = true
				}
			}
			if !classesFound {
				t.Fatal("missing dropdown class list")
			}
			charts := readDashboardCharts(t, path)
			wantCharts := 3
			if tc.hasHistory {
				wantCharts = 4
			}
			if len(charts) != wantCharts {
				t.Fatalf("expected %d charts, got %d", wantCharts, len(charts))
			}
			for i := range 2 {
				if len(charts[i].Chart.PlotArea.Donuts) != 1 {
					t.Fatalf("chart %d expected donut, got %+v", i, charts[i])
				}
			}
			shipped := charts[0].Chart.PlotArea.Donuts[0].Series[0]
			dollars := charts[1].Chart.PlotArea.Donuts[0].Series[0]
			comparison := charts[2].Chart.PlotArea.Bars[0].Series[0]
			assertDashboardRange(t, f, shipped.Categories, paddedDashboardValues("B - Beta", "A - Alpha"))
			assertDashboardRange(t, f, shipped.Values, paddedDashboardNumbers("20", "10"))
			assertDashboardRange(t, f, dollars.Categories, paddedDashboardValues("A - Alpha", "B - Beta"))
			assertDashboardRange(t, f, dollars.Values, paddedDashboardNumbers("300", "100"))
			assertDashboardRange(t, f, comparison.Values, tc.wantSales)
			wantLabels := []string{"YTD", "Previous Year (Full Year)"}
			if tc.hasHistory {
				wantLabels = []string{"2026 (completed months)", "2025 (same months)"}
			}
			assertDashboardRange(t, f, comparison.Categories, wantLabels)

			// Selecting one class must not include another class's items.
			setDashboardTestSelector(t, f, "B6", "Cards")
			assertDashboardRange(t, f, shipped.Categories, paddedDashboardValues("A - Alpha"))
			assertDashboardRange(t, f, shipped.Values, paddedDashboardNumbers("10"))
			assertDashboardRange(t, f, dollars.Values, paddedDashboardNumbers("300", "100"))
			assertDashboardRange(t, f, comparison.Values, tc.wantSales)

			setDashboardTestSelector(t, f, "L6", "CardsX")
			assertDashboardRange(t, f, dollars.Categories, paddedDashboardValues("B - Beta"))
			assertDashboardRange(t, f, dollars.Values, paddedDashboardNumbers("100"))
			assertDashboardRange(t, f, shipped.Values, paddedDashboardNumbers("10"))

			setDashboardTestSelector(t, f, "B50", "CardsX")
			assertDashboardRange(t, f, comparison.Values, tc.wantFilteredSales)
			if tc.hasHistory {
				monthly := charts[3].Chart.PlotArea.Bars[0].Series
				if len(monthly) != 2 {
					t.Fatalf("expected two monthly series, got %d", len(monthly))
				}
				for year, expected := range []string{"2026", "2025"} {
					parts := strings.SplitN(monthly[year].Name, "!", 2)
					if len(parts) != 2 {
						t.Fatalf("monthly series %d must reference a year cell, got %q", year, monthly[year].Name)
					}
					got, err := f.GetCellValue(strings.Trim(parts[0], "'"), parts[1])
					if err != nil || got != expected {
						t.Errorf("monthly series %d: expected legend %q, got %q from %s (error %v)", year, expected, got, monthly[year].Name, err)
					}
				}
				assertDashboardRange(t, f, monthly[0].Categories, []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"})
				assertDashboardRange(t, f, monthly[0].Values, []string{"30", "0", "0", "0", "0", "0", "0", "80", "1000", "0", "0", "0"})
				assertDashboardRange(t, f, monthly[1].Values, []string{"21", "0", "0", "0", "0", "0", "0", "56", "1000", "0", "0", "0"})
				setDashboardTestSelector(t, f, "L50", "Cards")
				assertDashboardRange(t, f, monthly[0].Values, []string{"10", "0", "0", "0", "0", "0", "0", "70", "900", "0", "0", "0"})
				assertDashboardRange(t, f, monthly[1].Values, []string{"7", "0", "0", "0", "0", "0", "0", "49", "800", "0", "0", "0"})
				assertDashboardRange(t, f, comparison.Values, tc.wantFilteredSales)
			}
		})
	}
}

// TestDashboardEmpty verifies that a workbook without items still has usable
// All selectors and zero-valued chart ranges, rather than broken references.
func TestDashboardEmpty(t *testing.T) {
	f := newProductLineWorkbook()
	t.Cleanup(func() { _ = f.Close() })
	if err := writeDashboardSheet(f, nil, false, time.Time{}); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "empty.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	written, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = written.Close() })
	charts := readDashboardCharts(t, path)
	assertDashboardRange(t, written, charts[0].Chart.PlotArea.Donuts[0].Series[0].Values, paddedDashboardNumbers())
	assertDashboardRange(t, written, charts[2].Chart.PlotArea.Bars[0].Series[0].Values, []string{"0", "0"})
}

// readDashboardCharts decodes the saved chart parts in creation order, returning
// chart series or failing with the specific unreadable archive member.
func readDashboardCharts(t *testing.T, path string) []dashboardChartXML {
	t.Helper()
	archive, err := zip.OpenReader(path)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = archive.Close() })
	var charts []dashboardChartXML
	for index := 1; ; index++ {
		name := fmt.Sprintf("xl/charts/chart%d.xml", index)
		reader, err := archive.Open(name)
		if err != nil {
			break
		}
		contents, err := io.ReadAll(reader)
		_ = reader.Close()
		if err != nil {
			t.Fatalf("read %s: %v", name, err)
		}
		var chart dashboardChartXML
		if err := xml.Unmarshal(contents, &chart); err != nil {
			t.Fatalf("decode %s: %v", name, err)
		}
		charts = append(charts, chart)
	}
	return charts
}

// assertDashboardRange calculates a saved chart's one-column range and compares
// every plotted value with independent expected values, with cell diagnostics.
func assertDashboardRange(t *testing.T, f *excelize.File, reference string, want []string) {
	t.Helper()
	parts := strings.SplitN(reference, "!", 2)
	if len(parts) != 2 {
		t.Fatalf("invalid chart reference %q", reference)
	}
	sheet := strings.Trim(parts[0], "'")
	bounds := strings.Split(strings.ReplaceAll(parts[1], "$", ""), ":")
	if len(bounds) != 2 {
		t.Fatalf("invalid chart range %q", reference)
	}
	col, first, err := excelize.CellNameToCoordinates(bounds[0])
	if err != nil {
		t.Fatal(err)
	}
	lastCol, last, err := excelize.CellNameToCoordinates(bounds[1])
	if err != nil {
		t.Fatal(err)
	}
	if lastCol != col || last-first+1 != len(want) {
		t.Fatalf("range %s expected %d values, got columns %d-%d rows %d-%d", reference, len(want), col, lastCol, first, last)
	}
	for index, expected := range want {
		cell, _ := excelize.CoordinatesToCellName(col, first+index)
		got, err := f.CalcCellValue(sheet, cell)
		if err != nil || got != expected {
			t.Errorf("%s!%s expected %q, got %q error=%v", sheet, cell, expected, got, err)
		}
	}
}

// setDashboardTestSelector changes one chart's class as a user would in Excel.
func setDashboardTestSelector(t *testing.T, f *excelize.File, cell, class string) {
	t.Helper()
	if err := f.SetCellValue("Dashboard", cell, class); err != nil {
		t.Fatalf("set %s=%q: %v", cell, class, err)
	}
}

// paddedDashboardValues returns independent fifteen-slice label expectations.
func paddedDashboardValues(values ...string) []string {
	result := make([]string, 15)
	copy(result, values)
	return result
}

// paddedDashboardNumbers returns independent fifteen-slice numeric expectations.
func paddedDashboardNumbers(values ...string) []string {
	result := make([]string, 15)
	for index := range result {
		result[index] = "0"
	}
	copy(result, values)
	return result
}
