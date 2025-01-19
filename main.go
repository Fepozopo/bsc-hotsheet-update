package main

import (
	"fmt"
	"time"

	helpers "github.com/Fepozopo/bsc-hotsheet-update/helpers"
)

func main() {
	startTime := time.Now()

	logger, logFile, err := helpers.CreateLogger("main", "ERROR")
	if err != nil {
		logger.Printf("failed to create log file: %v", err)
		return
	}
	defer logFile.Close()

	// Start the main loop
	for {
		var hotsheet string
		fmt.Print("Which hotsheet do you want to update? (smd, bsc, 21c, exit): ")
		_, err := fmt.Scanln(&hotsheet)
		if err != nil {
			logger.Printf("failed to read input: %v", err)
			return
		}

		switch hotsheet {
		case "smd":
			err = helpers.CaseSMD()
			if err != nil {
				logger.Printf("failed to update SMD hotsheet: %v", err)
				return
			}
			fmt.Printf("Done!\nElapsed time: %v\n", time.Since(startTime))
			return

		case "bsc":
			err = helpers.CaseBSC()
			if err != nil {
				logger.Printf("failed to update BSC hotsheet: %v", err)
				return
			}
			fmt.Printf("Done!\nElapsed time: %v\n", time.Since(startTime))
			return

		case "21c":
			err = helpers.Case21C()
			if err != nil {
				logger.Printf("failed to update 2021co hotsheet: %v", err)
				return
			}
			fmt.Printf("Done!\nElapsed time: %v\n", time.Since(startTime))
			return

		case "exit":
			return
		default:
			fmt.Println("Invalid input. Please enter 'smd', 'bsc', '21c', or 'exit' (case sensitive).")
		}
	}
}
