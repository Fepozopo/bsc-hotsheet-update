# Makefile for building this Go program with CGO for multiple OS targets

# Alternatively, you can use zig for cross-compilation for all OS targets 
# CC="zig cc -target x86_64-linux"
# CC="zig cc -target x86_64-macos"

# Default architecture
GOARCH ?= amd64

# Base binary name (without extension)
BINARY_BASE := hotsheet-updater

# Create folder function
define MAKE_BIN_DIR
	@mkdir -p bin
endef

# Build for Windows
windows:
	@echo "Building for Windows with GOOS=windows, GOARCH=$(GOARCH)..."
	$(call MAKE_BIN_DIR)
	@echo "Using zig cross compiler for Windows"
	@echo "Building with CC="zig cc -target x86_64-windows""
	CGO_ENABLED=1 GOOS=windows GOARCH=$(GOARCH) CC="zig cc -target x86_64-windows" \
		go build -o bin/$(BINARY_BASE).exe .

# Build for Linux
linux:
	@echo "Building for Linux with GOOS=linux, GOARCH=$(GOARCH)..."
	$(call MAKE_BIN_DIR)
	@echo "Using gcc for Linux"
	@echo "Building with CC=gcc"
	CGO_ENABLED=1 GOOS=linux GOARCH=$(GOARCH) CC=gcc \
		go build -o bin/$(BINARY_BASE) .

# Build for macOS (Intel)
macos:
	@echo "Building for macOS (Intel) with GOOS=darwin, GOARCH=$(GOARCH)..."
	$(call MAKE_BIN_DIR)
	@echo "Using clang for macOS"
	@echo "Building with CC=clang"
	CGO_ENABLED=1 GOOS=darwin GOARCH=$(GOARCH) CC=clang \
		go build -o bin/$(BINARY_BASE) .

# Build for macOS (ARM)
macos-arm:
	@echo "Building for macOS (ARM) with GOOS=darwin, GOARCH=arm64..."
	$(call MAKE_BIN_DIR)
	@echo "Using clang for macOS (ARM)"
	@echo "Building with CC=clang"
	CGO_ENABLED=1 GOOS=darwin GOARCH=arm64 CC=clang \
		go build -o bin/$(BINARY_BASE) .

# Clean target to remove generated binaries and bin folder if needed
clean:
	@echo "Cleaning generated binaries..."
	@rm -rf bin

.PHONY: windows linux macos macos-arm clean