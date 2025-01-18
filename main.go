package main

import (
	"fmt"
	"time"
)

func main() {
	startTime := time.Now()

	logger, err := handlerCreateLogger()
	if err != nil {
		logger.Printf("ERROR: failed to create log file: %v", err)
		return
	}

	// Start the main loop
	for {
		var hotsheet string
		fmt.Print("Which hotsheet do you want to update? (smd, bsc, 21c, exit): ")
		fmt.Scanln(&hotsheet)

		switch hotsheet {
		case "smd":
			err = handlerCaseSMD()
			if err != nil {
				logger.Printf("ERROR: failed to update SMD hotsheet: %v", err)
				return
			}
			fmt.Printf("Done!\nElapsed time: %v\n", time.Since(startTime))
			return

		case "bsc":
			err = handlerCaseBSC()
			if err != nil {
				logger.Printf("ERROR: failed to update BSC hotsheet: %v", err)
				return
			}
			fmt.Printf("Done!\nElapsed time: %v\n", time.Since(startTime))
			return

		case "21c":
			err = handlerCase21C()
			if err != nil {
				logger.Printf("ERROR: failed to update 2021co hotsheet: %v", err)
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
