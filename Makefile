# Makefile for building this Go program with CGO for multiple OS targets

# Base binary name (without extension)
BINARY_BASE := hotsheet_updater

# Create folder function
define MAKE_BIN_DIR
	@mkdir -p bin
endef

# Build for Windows (AMD64)
windows-amd64:
	@echo "Building for Windows with GOOS=windows, GOARCH=amd64..."
	$(call MAKE_BIN_DIR)
	@echo "Using zig cross compiler for Windows"
	@echo "Building with CC="zig cc -target x86_64-windows""
	CGO_ENABLED=1 GOOS=windows GOARCH=amd64 CC="zig cc -target x86_64-windows" \
		go build -o bin/$(BINARY_BASE)_windows-amd64.exe .

# Build for Windows (ARM64)
windows-arm64:
	@echo "Building for Windows (ARM) with GOOS=windows, GOARCH=arm64..."
	$(call MAKE_BIN_DIR)
	@echo "Using zig cross compiler for Windows"
	@echo "Building with CC="zig cc -target aarch64-windows""
	CGO_ENABLED=1 GOOS=windows GOARCH=arm64 CC="zig cc -target aarch64-windows" \
		go build -o bin/$(BINARY_BASE)-windows-arm64.exe .

# Build for Linux (AMD64)
linux-amd64:
	@echo "Building for Linux with GOOS=linux, GOARCH=amd64..."
	$(call MAKE_BIN_DIR)
	@echo "Using gcc for Linux"
	@echo "Building with CC="zig cc -target x86_64-linux""
	CGO_ENABLED=1 GOOS=linux GOARCH=amd64 CC="zig cc -target x86_64-linux" \
		go build -o bin/$(BINARY_BASE)-linux-amd64 .

# Build for Linux (ARM64)
linux-arm64:
	@echo "Building for Linux (ARM) with GOOS=linux, GOARCH=arm64..."
	$(call MAKE_BIN_DIR)
	@echo "Using gcc for Linux (ARM)"
	@echo "Building with CC="zig cc -target aarch64-linux""
	CGO_ENABLED=1 GOOS=linux GOARCH=arm64 CC="zig cc -target aarch64-linux" \
		go build -o bin/$(BINARY_BASE)-linux-arm64 .

# Build for darwin (ARM64)
darwin-arm64:
	@echo "Building for macOS (ARM) with GOOS=darwin, GOARCH=arm64..."
	$(call MAKE_BIN_DIR)
	@echo "Using clang for macOS (ARM)"
	@echo "Building with CC=clang"
	CGO_ENABLED=1 GOOS=darwin GOARCH=arm64 CC=clang \
		go build -o bin/$(BINARY_BASE)-darwin-arm64 .

# Build all targets
all: windows-amd64 windows-arm64 linux-amd64 linux-arm64 darwin-arm64

# Clean target to remove generated binaries and bin folder if needed
clean:
	@echo "Cleaning generated binaries..."
	@rm -rf bin

.PHONY: windows-amd64 windows-arm64 linux-amd64 linux-arm64 darwin-arm64 all clean
