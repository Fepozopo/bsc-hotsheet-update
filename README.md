# Hotsheet Updater

## Description

This is a Go program that updates Excel documents with data from an 'Item Listing With Sales History' report pulled from Sage 100. It supports updating multiple product lines (21c, BSC, BJP, SMD) and their respective sheets (e.g., Everyday, Winter Holiday, Spring Holiday, A2 Notecards) in one run. The program automatically matches and updates PO numbers and quantities for each SKU, and provides a console progress bar for feedback during updates. All operations generate detailed logs in a dedicated `logs-bsc` folder in the system temp directory, with robust error handling for easier troubleshooting. The GUI allows you to select and update all product lines at once.

## Motivation

I made this for my boss, who used to manually go through the reports and update the hotsheet himself—a process that took hours per product line. Now, with support for multiple product lines and automated PO updates, the program saves significant time and reduces errors. The progress bar and detailed logs make it easier to track updates and troubleshoot any issues quickly.

## Prerequisites

Before you begin, ensure you have the following installed on your system:

- Go programming language (version 1.16 or later)
- A C compiler for your target platform (e.g., GCC for Linux, Clang for macOS, or MinGW for Windows) because this program uses the Fyne GUI toolkit, which requires C bindings.
- You can optionally use zig as a C cross-compiler for all OS targets. The Makefile uses zig for Windows and Linux by default, but you can also use it for macOS if you aren't on that platform.

## Quick Start

1. Clone the repository and navigate to the project root in a terminal.
2. Run `make <target>` to build and run the program. Replace `<target>` with one of the following targets: `windows-x86_64`, `windows-arm`, `linux-x86_64`, `linux-arm`, `macos-arm`.

## Usage

1. You can modify the values in each case to adjust which sheets and columns contain each value in the hotsheet. The updater supports multiple product lines and sheets, and will automatically match and update PO numbers and quantities for each SKU from the PO report.
```go
everyday := Update{
	Hotsheet:          fileHotsheetNew,
	Sheet:             "EVERYDAY",
	InventoryReport:   inventoryReport,
	POReport:          POReport,
	SkuCol:            "C",
	OnHandCol:         "D",
	OnPOCol1:          "E",
	OnPOCol2:          "G",
	OnPOCol3:          "I",
	OnPOColTotal:      "K",
	OnSOBOCol:         "L",
	YtdSoldIssuedCol:  "Q",
	PONumCol1:         "F",
	PONumCol2:         "H",
	PONumCol3:         "J",
	AverageMonthlyCol: "R",
}
```
2. The program will open a GUI window to select the product line(s) to update and the files to update. You can select and update all product lines at once.
3. Select the Excel document(s) you want to update, the stock report file, and the sales report file.
4. The program will create new updated Excel documents with the data from the two reports, showing a progress bar in the console during the update.
5. All logs are stored in a 'logs-bsc' folder in the temp directory after each update, making troubleshooting easier if any issues arise.

## Building the Program

To build the program, I've included a Makefile. You can run `make <target>` to build the program for different platforms. You can also run `make clean` to remove the `bin` folder that contains the compiled binaries.
The targets are:
```bash
make windows-x86_64
make windows-arm
make linux-x86_64
make linux-arm
make macos-arm
```
You can also run `make all` to build all the targets or `make clean` to remove the binaries and bin folder.

These commands will build the program for the specified platform and output the binary to the `bin` folder.

## 🤝 Contributing

### Clone the repo

```bash
git clone https://github.com/Fepozopo/bsc-hotsheet-update
cd bsc-hotsheet-update
```

### Submit a pull request

Sorry, I'm not accepting pull requests at this time.
