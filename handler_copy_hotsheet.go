package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"time"

	"github.com/xuri/excelize/v2"
)

func handlerCopyHotsheet(product, hotsheet string) (string, error) {
	// Get the current date
	currentDate := time.Now().Format("2006-01-02")

	// Create a new file path
	logFilePath := fmt.Sprintf("./logs/handlerCopyHotsheet_%v.log", currentDate)

	// Create or open the log file
	logFile, err := os.OpenFile(logFilePath, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	if err != nil {
		return "", fmt.Errorf("error creating or opening log file: %w", err)
	}
	defer logFile.Close()

	// Set the log output to the log file
	log.SetOutput(logFile)

	// Create a logger that writes to the log file
	logger := log.New(logFile, "INFO: ", log.Ldate|log.Ltime|log.Lshortfile)

	// Open the hotsheet workbook
	wbHotsheet, err := excelize.OpenFile(hotsheet)
	if err != nil {
		logger.Printf("failed to open hotsheet file %s: %v", hotsheet, err)
		return "", fmt.Errorf("failed to open hotsheet file %s: %w", hotsheet, err)
	}
	defer wbHotsheet.Close()

	// Create a new file path
	fileName := fmt.Sprintf("%s_Hotsheet_%v.xlsx", product, currentDate)
	baseDir := filepath.Dir(hotsheet)
	filePath := filepath.Join(baseDir, fileName)

	// Copy the hotsheet with a different name
	err = wbHotsheet.SaveAs(filePath)
	if err != nil {
		logger.Printf("failed to save hotsheet file %s: %v", filePath, err)
		return "", fmt.Errorf("failed to save hotsheet file %s: %w", filePath, err)
	}

	return filePath, nil
}
