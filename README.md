# Hotsheet Updater

## Description

This is a Go program that updates an Excel document with data from a stock report and a sales report that's pulled from Sage 100. It is designed to be run on a schedule to keep the Excel document up to date.

## Motivation

I made this for my boss. He used to manually go through the reports and update the hotsheet himself. This would take him a few hours per product line. When I found this out, I couldn't believe it and made this to help him out.

## Prerequisites

Before you begin, ensure you have the following installed on your system:

- Go programming language (version 1.16 or later)
- A C compiler for your target platform (e.g., GCC for Linux, Clang for macOS, or MinGW for Windows) because this program uses the Fyne GUI toolkit, which requires C bindings.

## Quick Start

1. Clone the repository and navigate to the project root in a terminal.
2. Run `go build && ./bsc-hotsheet-update` to build and run the program.

## Usage

1. You can modify the values in each case to adjust which sheets and columns contain each value. The comments above correspond to each input. The letters are columns in the hotseet.
``` go
// HOTSHEET | SHEET | REPORT | SKU | ON HAND | ON PO | ON SO/BO
stock := UpdateStock{fileHotsheetNew, "EVERYDAY", fileStockReport, "E", "F", "I", "K"}
stockHoliday := UpdateStock{fileHotsheetNew, "HOLIDAY", fileStockReport, "C", "D", "F", "H"}
// HOTSHEET | SHEET | REPORT | SKU | YTD
sales := UpdateSales{fileHotsheetNew, "EVERYDAY", fileSalesReport, "E", "P"}
salesHoliday := UpdateSales{fileHotsheetNew, "HOLIDAY", fileSalesReport, "C", "N"}
```
2. The program will prompt you to select the product line and the files to update. Select the Excel document you want to update, the stock report file, and the sales report file.
3. The program will create a new updated Excel document with the data from the two reports.
4. There will be a 'logs_bsc-hotsheet-update' folder created in the temp directory to store all the logs after each update. I found these helpful if my boss ran into an issue. He could just send me this folder and I could see what happened.

## 🤝 Contributing

### Clone the repo

```bash
git clone https://github.com/Fepozopo/bsc-hotsheet-update
cd bsc-hotsheet-update
```

### Build the project

```bash
go build
```

### Run the project

```bash
./bsc-hotsheet-update
```

### Submit a pull request

If you'd like to contribute, please fork the repository and open a pull request to the `main` branch.
