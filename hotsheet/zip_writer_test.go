package hotsheet

import (
	"archive/zip"
	"bytes"
	"errors"
	"io"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

// TestFastWorkbookZIPRoundTrip verifies generated archives preserve empty text,
// escaped Unicode, multi-block content, zero values, and formulas across saves.
// Parallel workbooks and repeated saves exercise independent pooled ZIP entries.
func TestFastWorkbookZIPRoundTrip(t *testing.T) {
	for _, tc := range []struct{ name, text string }{
		{"empty text", ""},
		{"escaped Unicode", "00123 & <cards> — café 🎉"},
		// Escaping expands this valid, sub-32K cell into more than 64 KiB of XML.
		{"multiple DEFLATE blocks", strings.Repeat("&<>", 9000)},
	} {
		t.Run(tc.name, func(t *testing.T) {
			t.Parallel()
			f := newProductLineWorkbook()
			t.Cleanup(func() { _ = f.Close() })
			for pass, text := range []string{tc.text, tc.text + " updated"} {
				if err := f.SetCellStr("All Products", "A1", text); err != nil {
					t.Fatalf("save %d: set text: %v", pass, err)
				}
				if err := f.SetCellValue("All Products", "A2", 0); err != nil {
					t.Fatalf("save %d: set zero: %v", pass, err)
				}
				if err := f.SetCellFormula("All Products", "B2", "SUM(A2,3)"); err != nil {
					t.Fatalf("save %d: set formula: %v", pass, err)
				}
				buffer, err := f.WriteToBuffer()
				if err != nil {
					t.Fatalf("save %d: create workbook archive: %v", pass, err)
				}
				archive, err := zip.NewReader(bytes.NewReader(buffer.Bytes()), int64(buffer.Len()))
				if err != nil {
					t.Fatalf("save %d: open workbook archive: %v", pass, err)
				}
				// Reading every member to EOF verifies its decompression and CRC,
				// including parts Excelize might not load when checking cell values.
				for _, member := range archive.File {
					reader, err := member.Open()
					if err != nil {
						t.Fatalf("save %d, part %s: open: %v", pass, member.Name, err)
					}
					_, readErr := io.Copy(io.Discard, reader)
					closeErr := reader.Close()
					if readErr != nil || closeErr != nil {
						t.Fatalf("save %d, part %s: expected valid ZIP data, got read error %v, close error %v", pass, member.Name, readErr, closeErr)
					}
				}
				written, err := excelize.OpenReader(bytes.NewReader(buffer.Bytes()))
				if err != nil {
					t.Fatalf("save %d: reopen workbook: %v", pass, err)
				}
				t.Cleanup(func() { _ = written.Close() })
				for _, check := range []struct{ cell, want string }{{"A1", text}, {"A2", "0"}} {
					got, err := written.GetCellValue("All Products", check.cell, excelize.Options{RawCellValue: true})
					if err != nil || got != check.want {
						t.Errorf("save %d, cell %s: expected %.80q (%d bytes), got %.80q (%d bytes), error %v", pass, check.cell, check.want, len(check.want), got, len(got), err)
					}
				}
				if got, err := written.GetCellFormula("All Products", "B2"); err != nil || got != "SUM(A2,3)" {
					t.Errorf("save %d, formula B2: expected SUM(A2,3), got %q (error %v)", pass, got, err)
				}
			}
		})
	}
}

// failingZIPOutput rejects all writes with err to exercise ZIP error propagation.
type failingZIPOutput struct {
	err error
}

// Write rejects p, returning zero accepted bytes and the configured error.
func (output failingZIPOutput) Write(p []byte) (int, error) {
	return 0, output.err
}

// TestFastZIPWriteFailure verifies a failed output reaches the ZIP caller rather
// than being swallowed when the compressor is closed and returned to its pool.
func TestFastZIPWriteFailure(t *testing.T) {
	t.Parallel()
	want := errors.New("ZIP output failed")
	archive := newFastZIPWriter(failingZIPOutput{err: want})
	entry, writeErr := archive.Create("data.xml")
	if writeErr == nil {
		_, writeErr = entry.Write([]byte("<data>inventory</data>"))
	}
	closeErr := archive.Close()
	if writeErr != nil && !errors.Is(writeErr, want) {
		t.Fatalf("expected output failure or buffered write, got write error %v", writeErr)
	}
	if closeErr != nil && !errors.Is(closeErr, want) {
		t.Fatalf("expected output failure or completed close, got close error %v", closeErr)
	}
	if !errors.Is(writeErr, want) && !errors.Is(closeErr, want) {
		t.Fatalf("expected wrapped output failure %v, got write error %v and close error %v", want, writeErr, closeErr)
	}
}
