package hotsheet

import (
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	dataInsightsSheetName             = "Data Insights"
	dataInsightsTitleText             = "Data Insights"
	dataInsightsLeftTableStartCol     = "A"
	dataInsightsRightTableStartCol    = "G"
	dataInsightsTitleRow              = 1
	dataInsightsSectionRow            = 3
	dataInsightsTableStartRow         = 5
	dataInsightsTableColumnCount      = 5
	dataInsightsOtherTableColumnCount = 5
)

const (
	dataInsightsColumnOccasion = iota
	dataInsightsColumnDate
	dataInsightsColumnYTD
	dataInsightsColumnPY
	dataInsightsColumnYoYDisplay
)

var dataInsightsTableColumnWidths = []float64{26, 31, 14, 14, 28}
var dataInsightsOtherTableColumnWidths = []float64{26, 31, 14, 14, 28}

// dataInsightsCell constructs an Excel cell name (e.g. "A5") from a column letter and row number.
func dataInsightsCell(col string, row int) string {
	return fmt.Sprintf("%s%d", col, row)
}

// dataInsightsTableEndCol calculates the last column of a Data Insights table based on the shared column count.
func dataInsightsTableEndCol(startCol string, columnCount int) (string, error) {
	cols, err := dataInsightsTableColumns(startCol, columnCount)
	if err != nil {
		return "", err
	}

	return cols[columnCount-1], nil
}

// writeDataInsightsSheet creates the "Data Insights" worksheet, with Counter Cards grouped by
// seasonal section on the left and Other Products grouped into one table per class on the right.
func writeDataInsightsSheet(f *excelize.File, entries []*inventoryEntry) error {
	now := time.Now()
	currentMonthsThrough := currentMonthsThrough(now)
	// Use the current month progress to annualize in-progress rows.
	rowsBySection := buildDataInsightsRows(entries, currentMonthsThrough, now)
	otherRowsByClass := buildOtherProductsDataInsightsRows(entries, currentMonthsThrough, now)

	sheetName := dataInsightsSheetName
	if _, err := f.NewSheet(sheetName); err != nil {
		return fmt.Errorf("failed to create %s sheet: %w", sheetName, err)
	}

	titleStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(),
		Font:      boldFont(14),
	})
	if err != nil {
		return fmt.Errorf("failed to create title style: %w", err)
	}

	sectionStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(),
		Border:    thinBlackBorder(),
		Fill:      patternFill(dataInsightsSectionFill),
		Font:      boldFont(),
	})
	if err != nil {
		return fmt.Errorf("failed to create section style: %w", err)
	}

	headerStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(),
		Border:    thinBlackBorder(),
		Fill:      patternFill(standardHeaderFill),
		Font:      boldFont(),
	})
	if err != nil {
		return fmt.Errorf("failed to create header style: %w", err)
	}

	dataStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(),
		Border:    thinBlackBorder(),
	})
	if err != nil {
		return fmt.Errorf("failed to create data style: %w", err)
	}

	currencyDataStyle, err := f.NewStyle(&excelize.Style{
		Alignment:    centeredAlignment(),
		Border:       thinBlackBorder(),
		CustomNumFmt: currencyNumFmt(),
	})
	if err != nil {
		return fmt.Errorf("failed to create currency data style: %w", err)
	}

	totalStyle, err := f.NewStyle(&excelize.Style{
		Alignment: centeredAlignment(),
		Border:    thinBlackBorder(),
		Fill:      patternFill(dataInsightsTotalFill),
		Font:      boldFont(),
	})
	if err != nil {
		return fmt.Errorf("failed to create total style: %w", err)
	}

	currencyTotalStyle, err := f.NewStyle(&excelize.Style{
		Alignment:    centeredAlignment(),
		Border:       thinBlackBorder(),
		Fill:         patternFill(dataInsightsTotalFill),
		Font:         boldFont(),
		CustomNumFmt: currencyNumFmt(),
	})
	if err != nil {
		return fmt.Errorf("failed to create currency total style: %w", err)
	}

	if err := setDataInsightsTableWidths(f, sheetName, dataInsightsLeftTableStartCol, dataInsightsTableColumnWidths); err != nil {
		return err
	}
	if err := setDataInsightsTableWidths(f, sheetName, dataInsightsRightTableStartCol, dataInsightsOtherTableColumnWidths); err != nil {
		return err
	}

	leftEndCol, err := dataInsightsTableEndCol(dataInsightsLeftTableStartCol, dataInsightsTableColumnCount)
	if err != nil {
		return err
	}
	titleEndCol, err := dataInsightsTableEndCol(dataInsightsRightTableStartCol, dataInsightsOtherTableColumnCount)
	if err != nil {
		return err
	}

	// The title spans the full width of both tables to stay visually centered on the sheet.
	titleCell := dataInsightsCell(dataInsightsLeftTableStartCol, dataInsightsTitleRow)
	if err := f.SetCellValue(sheetName, titleCell, dataInsightsTitleText); err != nil {
		return fmt.Errorf("failed to set sheet title: %w", err)
	}
	if err := f.MergeCell(sheetName, titleCell, dataInsightsCell(titleEndCol, dataInsightsTitleRow)); err != nil {
		return fmt.Errorf("failed to merge title cells: %w", err)
	}
	if err := f.SetCellStyle(sheetName, titleCell, dataInsightsCell(titleEndCol, dataInsightsTitleRow), titleStyle); err != nil {
		return fmt.Errorf("failed to style title row: %w", err)
	}

	// The "Counter Cards" subtitle for the left-side table.
	counterCardsCell := dataInsightsCell(dataInsightsLeftTableStartCol, dataInsightsSectionRow)
	if err := f.SetCellValue(sheetName, counterCardsCell, "Counter Cards"); err != nil {
		return fmt.Errorf("failed to set counter cards subtitle: %w", err)
	}
	if err := f.MergeCell(sheetName, counterCardsCell, dataInsightsCell(leftEndCol, dataInsightsSectionRow)); err != nil {
		return fmt.Errorf("failed to merge counter cards subtitle: %w", err)
	}
	if err := f.SetCellStyle(sheetName, counterCardsCell, dataInsightsCell(leftEndCol, dataInsightsSectionRow), sectionStyle); err != nil {
		return fmt.Errorf("failed to style counter cards subtitle: %w", err)
	}

	// The "Other Products" subtitle for the right side table.
	otherProductsCell := dataInsightsCell(dataInsightsRightTableStartCol, dataInsightsSectionRow)
	if err := f.SetCellValue(sheetName, otherProductsCell, "Other Products"); err != nil {
		return fmt.Errorf("failed to set other products subtitle: %w", err)
	}
	if err := f.MergeCell(sheetName, otherProductsCell, dataInsightsCell(titleEndCol, dataInsightsSectionRow)); err != nil {
		return fmt.Errorf("failed to merge other products subtitle: %w", err)
	}
	if err := f.SetCellStyle(sheetName, otherProductsCell, dataInsightsCell(titleEndCol, dataInsightsSectionRow), sectionStyle); err != nil {
		return fmt.Errorf("failed to style other products subtitle: %w", err)
	}

	leftSections := []dataInsightsSection{
		{
			Name:               "Spring",
			Headers:            []string{"Occasion", "Date", "YTD Sales", "PY Sales", "Status / Projected YoY"},
			Rows:               rowsBySection["Spring"],
			CurrencyStartIndex: dataInsightsColumnYTD,
			CurrencyEndIndex:   dataInsightsColumnPY,
			RenderRow: func(row dataInsightsRow) []interface{} {
				return []interface{}{row.Occasion, row.Date, row.DollarSoldYTD, row.DollarSoldPY, row.YoYDisplay}
			},
			RenderTotal: func(totalYTD, totalPY, totalProjected float64, rows []dataInsightsRow) []interface{} {
				return []interface{}{"Total", "", totalYTD, totalPY, dataInsightsSeasonTotalYoYDisplay(totalProjected, totalPY, rows)}
			},
		},
		{
			Name:               "Winter",
			Headers:            []string{"Occasion", "Date", "YTD Sales", "PY Sales", "Status / Projected YoY"},
			Rows:               rowsBySection["Winter"],
			CurrencyStartIndex: dataInsightsColumnYTD,
			CurrencyEndIndex:   dataInsightsColumnPY,
			RenderRow: func(row dataInsightsRow) []interface{} {
				return []interface{}{row.Occasion, row.Date, row.DollarSoldYTD, row.DollarSoldPY, row.YoYDisplay}
			},
			RenderTotal: func(totalYTD, totalPY, totalProjected float64, rows []dataInsightsRow) []interface{} {
				return []interface{}{"Total", "", totalYTD, totalPY, dataInsightsSeasonTotalYoYDisplay(totalProjected, totalPY, rows)}
			},
		},
		{
			Name:               "Everyday",
			Headers:            []string{"Occasion", "Date", "YTD Sales", "PY Sales", "Projected YoY"},
			Rows:               rowsBySection["Everyday"],
			CurrencyStartIndex: dataInsightsColumnYTD,
			CurrencyEndIndex:   dataInsightsColumnPY,
			RenderRow: func(row dataInsightsRow) []interface{} {
				return []interface{}{row.Occasion, row.Date, row.DollarSoldYTD, row.DollarSoldPY, row.YoYDisplay}
			},
			RenderTotal: func(totalYTD, totalPY, totalProjected float64, rows []dataInsightsRow) []interface{} {
				return []interface{}{"Total", "", totalYTD, totalPY, formatYoYFromProjectedSales(totalProjected, totalPY)}
			},
		},
	}

	rowNum := dataInsightsTableStartRow
	for idx, section := range leftSections {
		if idx > 0 {
			rowNum++
		}
		nextRow, err := writeDataInsightsSectionTable(f, sheetName, dataInsightsLeftTableStartCol, rowNum, section, sectionStyle, headerStyle, dataStyle, currencyDataStyle, totalStyle, currencyTotalStyle)
		if err != nil {
			return err
		}
		rowNum = nextRow
	}

	rightSections := buildOtherProductsDataInsightsSections(otherRowsByClass)

	rowNum = dataInsightsTableStartRow
	for idx, section := range rightSections {
		if idx > 0 {
			rowNum++
		}
		nextRow, err := writeDataInsightsSectionTable(f, sheetName, dataInsightsRightTableStartCol, rowNum, section, sectionStyle, headerStyle, dataStyle, currencyDataStyle, totalStyle, currencyTotalStyle)
		if err != nil {
			return err
		}
		rowNum = nextRow
	}

	return nil
}

// dataInsightsRowStyleIDs stores the two Excel style IDs needed to render one detail row.
// The full-row style ID and the currency-range style ID are kept together so callers can
// swap fills on a per-row basis without accidentally dropping the number format on YTD/PY cells.
type dataInsightsRowStyleIDs struct {
	RowStyleID      int
	CurrencyStyleID int
}

// dataInsightsSection captures all the information needed to render one section of the Data Insights sheet, including
// the section name, headers, rows, any optional per-row style IDs, which columns are currency values, and how to
// render the row and total values.
type dataInsightsSection struct {
	Name               string
	Headers            []string
	Rows               []dataInsightsRow
	RowStyleIDs        []dataInsightsRowStyleIDs
	CurrencyStartIndex int
	CurrencyEndIndex   int
	RenderRow          func(dataInsightsRow) []interface{}
	RenderTotal        func(totalYTD, totalPY, totalProjected float64, rows []dataInsightsRow) []interface{}
}

// buildOtherProductsDataInsightsSections converts the class-keyed Other Products rows into the
// sheet sections rendered on the right side of Data Insights.
func buildOtherProductsDataInsightsSections(rowsByClass map[string][]dataInsightsRow) []dataInsightsSection {
	classNames := sortedOtherProductsDataInsightsClasses(rowsByClass)
	sections := make([]dataInsightsSection, 0, len(classNames))
	for _, className := range classNames {
		classRows := rowsByClass[className]
		// Each class owns its own table now, so repeating the class value in every detail row would
		// waste space that is better spent on the occasion/date/projection columns.
		sections = append(sections, dataInsightsSection{
			Name:               className,
			Headers:            []string{"Occasion", "Date", "YTD Sales", "PY Sales", "Status / Projected YoY"},
			Rows:               classRows,
			CurrencyStartIndex: dataInsightsColumnYTD,
			CurrencyEndIndex:   dataInsightsColumnPY,
			RenderRow: func(row dataInsightsRow) []interface{} {
				return []interface{}{row.Occasion, row.Date, row.DollarSoldYTD, row.DollarSoldPY, row.YoYDisplay}
			},
			RenderTotal: func(totalYTD, totalPY, totalProjected float64, rows []dataInsightsRow) []interface{} {
				return []interface{}{"Total", "", totalYTD, totalPY, otherProductsClassTotalYoYDisplay(totalProjected, totalPY, rows)}
			},
		})
	}

	return sections
}

// writeDataInsightsSectionTable renders one Data Insights table, including the section title,
// headers, detail rows, and total row. The caller supplies the row formatter so the same
// layout code can be reused for both Counter Cards and Other Products, while optional
// RowStyleIDs let a caller override the default row styling on a row-by-row basis.
func writeDataInsightsSectionTable(f *excelize.File, sheetName, startCol string, startRow int, section dataInsightsSection, sectionStyle, headerStyle, dataStyle, currencyDataStyle, totalStyle, currencyTotalStyle int) (int, error) {
	cols, err := dataInsightsTableColumns(startCol, len(section.Headers))
	if err != nil {
		return 0, err
	}
	endCol := cols[len(cols)-1]

	sectionTitleCell := dataInsightsCell(startCol, startRow)
	sectionTitleEnd := dataInsightsCell(endCol, startRow)
	if err := f.SetCellValue(sheetName, sectionTitleCell, section.Name); err != nil {
		return 0, fmt.Errorf("failed to set section title %s: %w", section.Name, err)
	}
	if err := f.MergeCell(sheetName, sectionTitleCell, sectionTitleEnd); err != nil {
		return 0, fmt.Errorf("failed to merge section title %s: %w", section.Name, err)
	}
	if err := f.SetCellStyle(sheetName, sectionTitleCell, sectionTitleEnd, sectionStyle); err != nil {
		return 0, fmt.Errorf("failed to style section title %s: %w", section.Name, err)
	}

	rowNum := startRow + 1
	for colIdx, header := range section.Headers {
		cell := dataInsightsCell(cols[colIdx], rowNum)
		if err := f.SetCellValue(sheetName, cell, header); err != nil {
			return 0, fmt.Errorf("failed to set header %s: %w", header, err)
		}
	}
	if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[0], rowNum), dataInsightsCell(endCol, rowNum), headerStyle); err != nil {
		return 0, fmt.Errorf("failed to style header row for %s: %w", section.Name, err)
	}

	rowNum++
	totalYTD := 0.0
	totalPY := 0.0
	totalProjectedSales := 0.0
	if len(section.RowStyleIDs) > 0 && len(section.RowStyleIDs) != len(section.Rows) {
		return 0, fmt.Errorf("failed to render %s rows: got %d row style entries, want %d", section.Name, len(section.RowStyleIDs), len(section.Rows))
	}
	for rowIdx, row := range section.Rows {
		values := section.RenderRow(row)
		if len(values) != len(section.Headers) {
			return 0, fmt.Errorf("failed to render %s row: got %d values, want %d", section.Name, len(values), len(section.Headers))
		}
		for colIdx, value := range values {
			cell := dataInsightsCell(cols[colIdx], rowNum)
			if err := f.SetCellValue(sheetName, cell, value); err != nil {
				return 0, fmt.Errorf("failed to write %s row cell %s: %w", section.Name, cell, err)
			}
		}
		rowStyleID := dataStyle
		rowCurrencyStyleID := currencyDataStyle
		if len(section.RowStyleIDs) > 0 {
			rowStyleID = section.RowStyleIDs[rowIdx].RowStyleID
			rowCurrencyStyleID = section.RowStyleIDs[rowIdx].CurrencyStyleID
		}
		if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[0], rowNum), dataInsightsCell(endCol, rowNum), rowStyleID); err != nil {
			return 0, fmt.Errorf("failed to style %s data row: %w", section.Name, err)
		}
		if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[section.CurrencyStartIndex], rowNum), dataInsightsCell(cols[section.CurrencyEndIndex], rowNum), rowCurrencyStyleID); err != nil {
			return 0, fmt.Errorf("failed to style %s currency cells: %w", section.Name, err)
		}

		totalYTD += row.DollarSoldYTD
		totalPY += row.DollarSoldPY
		totalProjectedSales += row.ProjectedDollar
		rowNum++
	}

	// The total row keeps the actual YTD and PY sums in their respective columns; only the
	// rightmost column uses projected sales to derive the YoY percentage shown in the sheet.
	totalRowValues := section.RenderTotal(totalYTD, totalPY, totalProjectedSales, section.Rows)
	if len(totalRowValues) != len(section.Headers) {
		return 0, fmt.Errorf("failed to render %s total row: got %d values, want %d", section.Name, len(totalRowValues), len(section.Headers))
	}
	for colIdx, value := range totalRowValues {
		cell := dataInsightsCell(cols[colIdx], rowNum)
		if err := f.SetCellValue(sheetName, cell, value); err != nil {
			return 0, fmt.Errorf("failed to write %s total row cell %s: %w", section.Name, cell, err)
		}
	}
	if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[0], rowNum), dataInsightsCell(endCol, rowNum), totalStyle); err != nil {
		return 0, fmt.Errorf("failed to style total row for %s: %w", section.Name, err)
	}
	if err := f.SetCellStyle(sheetName, dataInsightsCell(cols[section.CurrencyStartIndex], rowNum), dataInsightsCell(cols[section.CurrencyEndIndex], rowNum), currencyTotalStyle); err != nil {
		return 0, fmt.Errorf("failed to style %s total currency cells: %w", section.Name, err)
	}

	return rowNum + 1, nil
}

// setDataInsightsTableWidths applies a shared width pattern starting at the given column.
func setDataInsightsTableWidths(f *excelize.File, sheetName, startCol string, widths []float64) error {
	cols, err := dataInsightsTableColumns(startCol, len(widths))
	if err != nil {
		return err
	}

	for idx, col := range cols {
		if err := f.SetColWidth(sheetName, col, col, widths[idx]); err != nil {
			return fmt.Errorf("failed to set width for column %s: %w", col, err)
		}
	}

	return nil
}

// dataInsightsTableColumns returns table columns starting at startCol.
func dataInsightsTableColumns(startCol string, columnCount int) ([]string, error) {
	startIdx, err := excelize.ColumnNameToNumber(startCol)
	if err != nil {
		return nil, fmt.Errorf("failed to resolve column %s: %w", startCol, err)
	}

	cols := make([]string, columnCount)
	for i := 0; i < len(cols); i++ {
		col, err := excelize.ColumnNumberToName(startIdx + i)
		if err != nil {
			return nil, fmt.Errorf("failed to resolve column %d: %w", startIdx+i, err)
		}
		cols[i] = col
	}

	return cols, nil
}
