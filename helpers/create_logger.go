package helpers

import (
	"fmt"
	"log"
	"os"
	"time"
)

func CreateLogger(name, flag string) (*log.Logger, *os.File, error) {
	// Get the current date
	currentDate := time.Now().Format("2006-01-02 15:04:05.000000000")

	// Create a new file path
	logFilePath := fmt.Sprintf("./logs/%v_%s.log", currentDate, name)

	// Create or open the log file
	logFile, err := os.OpenFile(logFilePath, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	if err != nil {
		return nil, nil, fmt.Errorf("error creating or opening log file: %w", err)
	}

	// Create a logger that writes to the log file
	logger := log.New(logFile, flag+": ", log.Ldate|log.Ltime|log.Lshortfile)

	return logger, logFile, nil
}
