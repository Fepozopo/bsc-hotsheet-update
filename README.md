# Hotsheet Updater

## Description

This is a Go program that updates an Excel document with data from an 'Item Listing With Sales History' report that's pulled from Sage 100. It is designed to be run on a schedule to keep the Excel document up to date.

## Motivation

I made this for my boss. He used to manually go through the reports and update the hotsheet himself. This would take him a few hours per product line. When I found this out, I couldn't believe it and made this to help him out.

## Prerequisites

Before you begin, ensure you have the following installed on your system:

- Go programming language (version 1.16 or later)
- A C compiler for your target platform (e.g., GCC for Linux, Clang for macOS, or MinGW for Windows) because this program uses the Fyne GUI toolkit, which requires C bindings.
- You can optionally use zig as a C cross-compiler for all OS targets. The Makefile uses zig for Windows and Linux by default, but you can also use it for macOS if you aren't on that platform.

## Quick Start

1. Clone the repository and navigate to the project root in a terminal.
2. Run `make <target>` to build and run the program. Replace `<target>` with one of the following targets: `windows-amd`, `windows-arm`, `linux-amd`, `linux-arm`, `macos-amd`, or `macos-arm`.

## Usage

1. You can modify the values in each case to adjust which sheets and columns contain each value in the hotsheet.
``` go
everyday := Update{
		Hotsheet:         fileHotsheetNew,
		Sheet:            "EVERYDAY",
		InventoryReport:  inventoryReport,
		POReport:         POReport,
		SkuCol:           "E",
		OnHandCol:        "F",
		OnPOCol:          "I",
		OnSOBOCol:        "K",
		YtdSoldIssuedCol: "P",
		PONumCol:         "J",
	}
```
2. The program will open a GUI window to select the product line to update and the files to update.
3. Select the Excel document you want to update, the stock report file, and the sales report file.
4. The program will create a new updated Excel document with the data from the two reports.
5. There will be a 'logs-bsc' folder created in the temp directory to store all the logs after each update. I found these helpful if my boss ran into an issue. He could just send me this folder and I could see what happened.

## Building the Program

To build the program, I've included a Makefile. You can run `make <target>` to build the program for different platforms. You can also run `make clean` to remove the `bin` folder that contains the compiled binaries.
The targets are:
```bash
make windows-amd
make windows-arm
make linux-amd
make linux-arm
make macos-amd
make macos-arm
```
You can also run `make all` to build all the targets.

These commands will build the program for the specified platform and output the binary to the `bin` folder.

## 🤝 Contributing

### Clone the repo

```bash
git clone https://github.com/Fepozopo/bsc-hotsheet-update
cd bsc-hotsheet-update
```

### Submit a pull request

Sorry, I'm not accepting pull requests at this time.
