package helpers

import (
	"fmt"
	"log"
	"os"
	"time"
)

func CreateLogger() (*log.Logger, error) {
	// Get the current date
	currentDate := time.Now().Format("2006-01-02 15:04:05.000000000")

	// Create a new file path
	logFilePath := fmt.Sprintf("./logs/%v.log", currentDate)

	// Create or open the log file
	logFile, err := os.OpenFile(logFilePath, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	if err != nil {
		return nil, fmt.Errorf("error creating or opening log file: %w", err)
	}
	defer logFile.Close()

	// Set the log output to the log file
	log.SetOutput(logFile)

	// Create a logger that writes to the log file
	logger := log.New(logFile, "INFO: ", log.Ldate|log.Ltime|log.Lshortfile)

	return logger, nil
}
