package helpers

import (
	"bufio"
	"fmt"
	"io"
	"log/slog"
	"os"
	"path/filepath"
	"strings"
	"time"
)

// CreateSlogLogger creates a buffered slog-based logger that writes JSON entries to a
// timestamped log file inside the system temp directory under "logs-bsc"
// It returns the created *slog.Logger and an io.Closer
// that must be called on shutdown to flush the buffer, sync and close the underlying file.
//
//   - name: logical name of the operation (e.g. "create", "main").
//   - product, occasion: optional components that are appended to the filename if provided.
//   - level: textual log level ("DEBUG", "INFO", "WARN", "ERROR") - case-insensitive.
//     Unrecognized values fall back to INFO.
//
// Important: callers must call Close() on the returned io.Closer (or otherwise ensure
// the buffer is flushed and the file closed) to avoid losing recently-buffered log entries.
func CreateSlogLogger(name, product, occasion, level string) (*slog.Logger, io.Closer, error) {
	// Timestamped filename
	currentDate := time.Now().Format("2006-01-02_150405.000000000")

	tempDir := os.TempDir()
	logDir := filepath.Join(tempDir, "logs-bsc")

	if err := os.MkdirAll(logDir, os.ModePerm); err != nil {
		return nil, nil, fmt.Errorf("error creating logs directory: %w", err)
	}

	var logFilePath string
	// Build filename components to avoid empty placeholder segments.
	base := fmt.Sprintf("%s_%s", currentDate, name)
	parts := []string{base}
	if product != "" {
		parts = append(parts, product)
	}
	if occasion != "" {
		parts = append(parts, occasion)
	}
	filename := strings.Join(parts, "-") + ".log"
	logFilePath = filepath.Join(logDir, filename)

	f, err := os.OpenFile(logFilePath, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0o644)
	if err != nil {
		return nil, nil, fmt.Errorf("error creating or opening log file: %w", err)
	}

	// Buffered writer to reduce syscalls. 64KB buffer is a reasonable default.
	bw := bufio.NewWriterSize(f, 64*1024)

	// Create slog handler writing JSON to the buffered writer.
	h := slog.NewJSONHandler(bw, &slog.HandlerOptions{
		Level: parseSlogLevel(level),
	})

	logger := slog.New(h)

	closer := &slogFileCloser{
		file: f,
		bw:   bw,
	}

	return logger, closer, nil
}

// slogFileCloser flushes the bufio.Writer, syncs and closes the underlying file.
type slogFileCloser struct {
	file *os.File
	bw   *bufio.Writer
}

func (c *slogFileCloser) Close() error {
	var firstErr error

	// helper to record the first error encountered
	setFirstErr := func(err error, prefix string) {
		if err != nil && firstErr == nil {
			firstErr = fmt.Errorf("%s: %w", prefix, err)
		}
	}

	// Flush buffered data first
	if c.bw != nil {
		if err := c.bw.Flush(); err != nil {
			setFirstErr(err, "buffer flush error")
		}
	}

	// Sync file contents to disk and close the file
	if c.file != nil {
		if err := c.file.Sync(); err != nil {
			setFirstErr(err, "file sync error")
		}
		if err := c.file.Close(); err != nil {
			setFirstErr(err, "file close error")
		}
	}

	return firstErr
}

// parseSlogLevel maps a textual level to slog.Level. Defaults to slog.LevelInfo for unknown levels.
func parseSlogLevel(level string) slog.Level {
	switch strings.ToUpper(strings.TrimSpace(level)) {
	case "DEBUG", "D":
		return slog.LevelDebug
	case "INFO", "I":
		return slog.LevelInfo
	case "WARN", "WARNING", "W":
		return slog.LevelWarn
	case "ERROR", "ERR", "E":
		return slog.LevelError
	default:
		return slog.LevelInfo
	}
}
