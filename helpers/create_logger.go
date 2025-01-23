package helpers

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"time"
)

// CreateLogger creates a logger that writes to a log file.
// It takes the logger name and a log flag as parameters.
func CreateLogger(name, product, occasion, flag string) (*log.Logger, *os.File, error) {
	// Get the current date
	currentDate := time.Now().Format("2006-01-02_150405.000000000")

	// Get the temporary directory
	tempDir := os.TempDir()

	// Define the logs directory path
	logDir := filepath.Join(tempDir, "logs_bsc-hotsheet-update")

	// Create a new file path
	var logFilePath string
	if product == "" && occasion == "" {
		logFilePath = filepath.Join(logDir, fmt.Sprintf("%s_%s.log", currentDate, name))
	} else {
		logFilePath = filepath.Join(logDir, fmt.Sprintf("%s_%s-%s-%s.log", currentDate, name, product, occasion))
	}

	// Create the logs directory if it does not exist
	if err := os.MkdirAll(logDir, os.ModePerm); err != nil {
		return nil, nil, fmt.Errorf("error creating logs directory: %w", err)
	}

	// Create or open the log file
	logFile, err := os.OpenFile(logFilePath, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	if err != nil {
		return nil, nil, fmt.Errorf("error creating or opening log file: %w", err)
	}

	// Create a logger that writes to the log file
	logger := log.New(logFile, flag+": ", log.Ldate|log.Ltime|log.Lshortfile)

	return logger, logFile, nil
}
