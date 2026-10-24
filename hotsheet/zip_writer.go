// Workbook ZIP output uses pooled level-1 DEFLATE compression.
package hotsheet

import (
	"archive/zip"
	"compress/flate"
	"io"
	"sync"

	"github.com/xuri/excelize/v2"
)

// fastZIPCompressors reuses compression buffers across ZIP entries and workbooks.
// Each checked-out compressor is owned by one entry until that entry closes.
var fastZIPCompressors sync.Pool

// newFastZIPWriter returns an Excelize-compatible archive writer targeting dst.
// Only this archive's DEFLATE compressor is overridden; global ZIP defaults are
// unchanged. Level 1 favors workbook generation speed over compressed size.
func newFastZIPWriter(dst io.Writer) excelize.ZipWriter {
	writer := zip.NewWriter(dst)
	writer.RegisterCompressor(zip.Deflate, newFastZIPCompressor)
	return writer
}

// newFastZIPCompressor returns a level-1 DEFLATE writer targeting dst, reusing
// buffers when available. It returns an error if compressor creation fails.
func newFastZIPCompressor(dst io.Writer) (io.WriteCloser, error) {
	writer, ok := fastZIPCompressors.Get().(*flate.Writer)
	if ok {
		writer.Reset(dst)
	} else {
		var err error
		writer, err = flate.NewWriter(dst, flate.BestSpeed)
		if err != nil {
			return nil, err
		}
	}
	return &pooledZIPCompressor{writer: writer}, nil
}

// pooledZIPCompressor returns its DEFLATE writer to the pool on Close. Like a
// ZIP entry writer, one instance must not be used concurrently by callers.
type pooledZIPCompressor struct {
	writer *flate.Writer
}

// Write compresses p, returning the accepted byte count and any output error.
// Writes after Close fail with io.ErrClosedPipe rather than touching a reused writer.
func (compressor *pooledZIPCompressor) Write(p []byte) (int, error) {
	if compressor.writer == nil {
		return 0, io.ErrClosedPipe
	}
	return compressor.writer.Write(p)
}

// Close flushes the entry and releases its compressor for reuse, returning any
// flush error. Repeated calls are harmless and never return a writer twice.
func (compressor *pooledZIPCompressor) Close() error {
	writer := compressor.writer
	if writer == nil {
		return nil
	}
	compressor.writer = nil
	err := writer.Close()
	// Reset releases the previous output buffer and clears error state, so the
	// pool neither retains a workbook archive nor contaminates the next entry.
	writer.Reset(io.Discard)
	fastZIPCompressors.Put(writer)
	return err
}
