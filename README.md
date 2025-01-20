# BSC Hotsheet Updater

This is a small Go program that updates an Excel document with data from a stock report and a sales report that's pulled from Sage 100. It is designed to be run on a schedule to keep the Excel document up to date.

## Setup

1. Clone the repository and navigate to the project root in a terminal.
2. Run `go get` to install all the required packages.
3. You can modify the values in each case to adjust which sheets and columns contain each value. It looks like this:
``` go
    // HOTSHEET | SHEET | REPORT | SKU | ON HAND | ON PO | ON SO/BO
	stock := UpdateStock{fileHotsheetNew, "EVERYDAY", fileStockReport, "E", "F", "I", "K"}
	stockHoliday := UpdateStock{fileHotsheetNew, "HOLIDAY", fileStockReport, "C", "D", "F", "H"}
	// HOTSHEET | SHEET | REPORT | SKU | YTD
	sales := UpdateSales{fileHotsheetNew, "EVERYDAY", fileSalesReport, "E", "P"}
	salesHoliday := UpdateSales{fileHotsheetNew, "HOLIDAY", fileSalesReport, "C", "N"}
```

## Usage

1. Run `go run .` to start the program.
2. The program will prompt you to select the files to update. Select the Excel document you want to update, the stock report file, and the sales report file.
3. The program will create a new updated Excel document with the data from the two reports.
4. There will be a 'logs' folder created in the current directory to store all the logs after each update.

## Development

1. Run `go build` to build the program.
2. Run `./bsc-hotsheet-update` to run the program.
