# Makefile for building this Go program for multiple OS targets without CGO

# Base binary name (without extension)
BINARY_BASE := hotsheet

# Create folder function
define MAKE_BIN_DIR
	@mkdir -p bin
endef

# Build for Windows (AMD64)
windows-amd64:
	@echo "Building for Windows with GOOS=windows, GOARCH=amd64 (CGO disabled)..."
	$(call MAKE_BIN_DIR)
	GOOS=windows GOARCH=amd64 \
		go build -o bin/$(BINARY_BASE)-windows-amd64.exe .

# Build for Windows (ARM64)
windows-arm64:
	@echo "Building for Windows (ARM) with GOOS=windows, GOARCH=arm64 (CGO disabled)..."
	$(call MAKE_BIN_DIR)
	GOOS=windows GOARCH=arm64 \
		go build -o bin/$(BINARY_BASE)-windows-arm64.exe .

# Build for Linux (AMD64)
linux-amd64:
	@echo "Building for Linux with GOOS=linux, GOARCH=amd64 (CGO disabled)..."
	$(call MAKE_BIN_DIR)
	GOOS=linux GOARCH=amd64 \
		go build -o bin/$(BINARY_BASE)-linux-amd64 .

# Build for Linux (ARM64)
linux-arm64:
	@echo "Building for Linux (ARM) with GOOS=linux, GOARCH=arm64 (CGO disabled)..."
	$(call MAKE_BIN_DIR)
	GOOS=linux GOARCH=arm64 \
		go build -o bin/$(BINARY_BASE)-linux-arm64 .

# Build for darwin (ARM64)
darwin-arm64:
	@echo "Building for macOS (ARM) with GOOS=darwin, GOARCH=arm64 (CGO disabled)..."
	$(call MAKE_BIN_DIR)
	GOOS=darwin GOARCH=arm64 \
		go build -o bin/$(BINARY_BASE)-darwin-arm64 .

# Build all targets
all: windows-amd64 windows-arm64 linux-amd64 linux-arm64 darwin-arm64

# Clean target to remove generated binaries and bin folder if needed
clean:
	@echo "Cleaning generated binaries..."
	@rm -rf bin

.PHONY: windows-amd64 windows-arm64 linux-amd64 linux-arm64 darwin-arm64 all clean
