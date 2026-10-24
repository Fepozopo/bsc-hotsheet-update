package hotsheet

import (
	"bytes"
	"encoding/json"
	"errors"
	"log/slog"
	"slices"
	"testing"
	"testing/synctest"
	"time"
)

// fakeReportReader replaces the workbook boundary with controlled parser calls.
// Each test configures only the methods its selected paths should invoke.
type fakeReportReader struct {
	history func(string, bool) (monthlyHistoryData, error)
	po      func(string) (poImportData, error)
}

// readHistory delegates the supplied path and issue selector to the test parser.
func (reader fakeReportReader) readHistory(path string, issue bool) (monthlyHistoryData, error) {
	return reader.history(path, issue)
}

// readPO delegates the supplied path to the test PO parser.
func (reader fakeReportReader) readPO(path string) (poImportData, error) {
	return reader.po(path)
}

// TestReportImportPriorityAndMerge verifies the two history readers start first,
// PO fills the first freed slot, and no merge occurs while readers remain active.
// Final history ordering and date anchoring must not depend on completion order.
func TestReportImportPriorityAndMerge(t *testing.T) {
	synctest.Test(t, func(t *testing.T) {
		gates := map[string]chan struct{}{}
		for _, name := range []string{"sales", "issue", "po"} {
			gates[name] = generationTestGate(t)
		}
		started := make(chan string, 3)
		item := &inventoryEntry{SKU: "SKU-1", SalesRecords: []salesRecord{{Year: 2023, Metric: "Quantity Sold"}}}
		salesDate := time.Date(2026, time.September, 25, 0, 0, 0, 0, time.UTC)
		reader := fakeReportReader{
			history: func(path string, issue bool) (monthlyHistoryData, error) {
				started <- path
				<-gates[path]
				if issue {
					if path != "issue" {
						t.Errorf("issue selector: expected path issue, got %s", path)
					}
					return monthlyHistoryData{
						runDate: time.Date(2026, time.September, 29, 0, 0, 0, 0, time.UTC),
						records: map[*inventoryEntry][]salesRecord{item: {{Year: 2025, Metric: "Quantity Issued", Periods: [12]float64{3}}}},
					}, nil
				}
				if path != "sales" {
					t.Errorf("sales selector: expected path sales, got %s", path)
				}
				return monthlyHistoryData{
					runDate: salesDate,
					records: map[*inventoryEntry][]salesRecord{item: {{Year: 2024, Metric: "Quantity Sold", Periods: [12]float64{2}}}},
				}, nil
			},
			po: func(path string) (poImportData, error) {
				started <- path
				<-gates[path]
				return poImportData{item: {{number: "123", quantity: 7}}}, nil
			},
		}
		done := make(chan struct{})
		var gotDate time.Time
		var gotErr error
		var progress []Progress
		go func() {
			gotDate, gotErr = mergeOptionalReports("po", "sales", "issue", reader, nil, func(update Progress) {
				progress = append(progress, update)
			})
			close(done)
		}()
		synctest.Wait()
		if len(started) != 2 {
			t.Fatalf("blocked history readers: expected exactly 2 starts, got %d", len(started))
		}
		initial := []string{<-started, <-started}
		slices.Sort(initial)
		if !slices.Equal(initial, []string{"issue", "sales"}) {
			t.Fatalf("expected sales and issue to start first, got %v", initial)
		}
		close(gates["issue"])
		synctest.Wait()
		if len(started) != 1 {
			t.Fatalf("issue completed with sales blocked: expected one PO start, got %d", len(started))
		}
		if got := <-started; got != "po" {
			t.Fatalf("issue freed a slot: expected PO, got %s", got)
		}
		close(gates["po"])
		synctest.Wait()
		if item.PONum1 != "" || !slices.Equal(item.SalesRecords, []salesRecord{{Year: 2023, Metric: "Quantity Sold"}}) {
			t.Errorf("sales still active: expected untouched inventory, got %+v", item)
		}
		select {
		case <-done:
			t.Fatal("returned while the sales reader was still active")
		default:
		}
		close(gates["sales"])
		<-done
		wantRecords := []salesRecord{
			{Year: 2023, Metric: "Quantity Sold"},
			{Year: 2024, Metric: "Quantity Sold", Periods: [12]float64{2}},
			{Year: 2025, Metric: "Quantity Issued", Periods: [12]float64{3}},
		}
		if gotErr != nil || !gotDate.Equal(salesDate) || !slices.Equal(item.SalesRecords, wantRecords) || item.PONum1 != "123" || item.OnPO1 != 7 {
			t.Errorf("completion order issue,PO,sales: expected sales date %v, records %v, PO 123/7 and no error; got date %v, item %+v, error %v", salesDate, wantRecords, gotDate, item, gotErr)
		}
		if len(progress) == 0 {
			t.Fatal("expected import progress callbacks, got none")
		}
		previous := 0
		for _, update := range progress {
			if update.Percent < previous || update.Percent >= 100 {
				t.Errorf("after %d%%: expected monotonic import progress below 100, got %+v", previous, update)
			}
			previous = update.Percent
		}
	})
}

// TestReportImportErrors verifies history-error precedence is independent of
// finish order, all readers drain on failure, and PO failures remain nonfatal
// and logged even when a history report also fails.
func TestReportImportErrors(t *testing.T) {
	salesFailure := errors.New("sales failed")
	issueFailure := errors.New("issue failed")
	poFailure := errors.New("PO failed")
	for _, tc := range []struct {
		name                      string
		salesErr, issueErr, poErr error
		wantErr                   error
	}{
		{"sales error takes precedence", salesFailure, issueFailure, poFailure, salesFailure},
		{"issue error stops generation", nil, issueFailure, nil, issueFailure},
		{"PO error is nonfatal", nil, nil, poFailure, nil},
	} {
		t.Run(tc.name, func(t *testing.T) {
			synctest.Test(t, func(t *testing.T) {
				salesGate, issueGate, poGate := generationTestGate(t), generationTestGate(t), generationTestGate(t)
				salesDate := time.Date(2026, time.September, 25, 0, 0, 0, 0, time.UTC)
				reader := fakeReportReader{
					history: func(_ string, issue bool) (monthlyHistoryData, error) {
						if issue {
							<-issueGate
							return monthlyHistoryData{}, tc.issueErr
						}
						<-salesGate
						if tc.salesErr != nil {
							return monthlyHistoryData{}, tc.salesErr
						}
						return monthlyHistoryData{runDate: salesDate}, nil
					},
					po: func(string) (poImportData, error) {
						<-poGate
						return nil, tc.poErr
					},
				}
				var logs bytes.Buffer
				logger := slog.New(slog.NewJSONHandler(&logs, nil))
				done := make(chan struct{})
				var gotDate time.Time
				var gotErr error
				go func() {
					gotDate, gotErr = mergeOptionalReports("po", "sales", "issue", reader, logger, nil)
					close(done)
				}()
				synctest.Wait()
				close(issueGate)
				synctest.Wait()
				close(poGate)
				synctest.Wait()
				select {
				case <-done:
					t.Fatal("returned before the sales reader finished")
				default:
				}
				close(salesGate)
				<-done
				if !errors.Is(gotErr, tc.wantErr) {
					t.Errorf("issue finished first: expected error %v, got %v", tc.wantErr, gotErr)
				}
				wantDate := salesDate
				if tc.wantErr != nil {
					wantDate = time.Time{}
				}
				if !gotDate.Equal(wantDate) {
					t.Errorf("history errors sales=%v issue=%v: expected date %v, got %v", tc.salesErr, tc.issueErr, wantDate, gotDate)
				}
				if tc.poErr == nil {
					if logs.Len() != 0 {
						t.Errorf("successful PO read: expected no error log, got %s", logs.String())
					}
				} else {
					var record map[string]any
					if err := json.NewDecoder(&logs).Decode(&record); err != nil {
						t.Fatalf("PO failure %v: expected a structured error log, got decode error %v", tc.poErr, err)
					}
					if record["level"] != "ERROR" || record["err"] != tc.poErr.Error() {
						t.Errorf("PO failure %v: expected ERROR with that cause, got %v", tc.poErr, record)
					}
				}
			})
		})
	}
}

// TestReportImportOptionalPaths verifies absent reports never invoke a parser,
// nil callbacks and loggers are accepted, and sales alone still supplies the date.
func TestReportImportOptionalPaths(t *testing.T) {
	for _, tc := range []struct {
		name, po, sales, issue string
		wantCalls              []string
		wantDate               bool
	}{
		{name: "no reports"},
		{name: "whitespace reports", po: " ", sales: "\t", issue: "\n"},
		{name: "PO only", po: "po", wantCalls: []string{"po"}},
		{name: "sales only", sales: "sales", wantCalls: []string{"sales"}, wantDate: true},
		{name: "sales and PO", po: "po", sales: "sales", wantCalls: []string{"po", "sales"}, wantDate: true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			calls := make(chan string, 3)
			date := time.Date(2026, time.September, 25, 0, 0, 0, 0, time.UTC)
			reader := fakeReportReader{
				history: func(path string, issue bool) (monthlyHistoryData, error) {
					calls <- path
					if issue {
						t.Errorf("paths %+v: expected no issue parser call, got %q", tc, path)
					}
					return monthlyHistoryData{runDate: date}, nil
				},
				po: func(path string) (poImportData, error) { calls <- path; return nil, nil },
			}
			gotDate, err := mergeOptionalReports(tc.po, tc.sales, tc.issue, reader, nil, nil)
			close(calls)
			var gotCalls []string
			for path := range calls {
				gotCalls = append(gotCalls, path)
			}
			slices.Sort(gotCalls)
			hasDate := !gotDate.IsZero()
			if err != nil || !slices.Equal(gotCalls, tc.wantCalls) || hasDate != tc.wantDate {
				t.Errorf("paths PO=%q sales=%q issue=%q: expected calls %v, date present=%v and no error; got calls %v, date %v, error %v", tc.po, tc.sales, tc.issue, tc.wantCalls, tc.wantDate, gotCalls, gotDate, err)
			}
		})
	}
}

// TestOptionalReadersKeepInventoryPrivate verifies real Excelize readers retain
// their parsed data separately until merging, including issue clamping and PO
// status handling. The inventory report remains authoritative for existing fields.
func TestOptionalReadersKeepInventoryPrivate(t *testing.T) {
	item := &inventoryEntry{SKU: "SKU-1", Description: "inventory description", YTDSold: 42, YTDIssued: 4}
	inventory := map[string]*inventoryEntry{"SKU-1": item}
	reader := excelReportReader{inventoryBySKU: inventory}
	sales, err := reader.readHistory(writeSalesHistoryFixture(t), false)
	if err != nil {
		t.Fatalf("sales fixture: expected successful parse, got %v", err)
	}
	issue, err := reader.readHistory(writeIssueHistoryFixture(t), true)
	if err != nil {
		t.Fatalf("issue fixture: expected successful parse, got %v", err)
	}
	po, err := reader.readPO(writePOFixture(t))
	if err != nil {
		t.Fatalf("PO fixture: expected successful parse, got %v", err)
	}
	if len(item.SalesRecords) != 0 || item.PONum1 != "" || item.PONum2 != "" || item.OnPO1 != 0 || item.OnPO2 != 0 || len(inventory) != 1 {
		t.Fatalf("before merging: expected unchanged inventory and no report-only SKUs, got %+v and %d entries", item, len(inventory))
	}
	po.merge()
	sales.merge()
	issue.merge()
	if len(item.SalesRecords) != 7 {
		t.Fatalf("after merging: expected 6 sales records followed by 1 issue record, got %v", item.SalesRecords)
	}
	if item.SalesRecords[0].Metric != "Quantity Sold" || item.SalesRecords[0].Periods[0] != 4 || item.SalesRecords[6].Metric != "Quantity Issued" || item.SalesRecords[6].Periods[0] != 3 || item.SalesRecords[6].Periods[1] != 0 {
		t.Errorf("after merging: expected sales first and clamped issue quantities last, got %v", item.SalesRecords)
	}
	if item.PONum1 != "123" || item.OnPO1 != 8 || item.PONum2 != "45" || item.OnPO2 != 7 || item.Description != "inventory description" || item.YTDSold != 42 || item.YTDIssued != 4 {
		t.Errorf("after merging: expected PO1=123/8, PO2=45/7 and original inventory description/YTD values, got %+v", item)
	}
}

// TestHistoryReaderEmptyInventory verifies nil maps, empty maps, and nil entries
// do not import history-only SKUs or dereference missing inventory entries.
func TestHistoryReaderEmptyInventory(t *testing.T) {
	path := writeSalesHistoryFixture(t)
	for _, tc := range []struct {
		name      string
		inventory map[string]*inventoryEntry
	}{
		{name: "nil inventory"},
		{name: "empty inventory", inventory: map[string]*inventoryEntry{}},
		{name: "nil item", inventory: map[string]*inventoryEntry{"SKU-1": nil}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			data, err := readMonthlyHistory(path, tc.inventory, nil, false)
			if err != nil || len(data.records) != 0 {
				t.Errorf("inventory %v: expected no matching records and no error, got %v and %v", tc.inventory, data.records, err)
			}
			data.merge()
		})
	}
}
