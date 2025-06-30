# Makefile for building this Go program with CGO for multiple OS targets

# Base binary name (without extension)
BINARY_BASE := hotsheet_updater

# Create folder function
define MAKE_BIN_DIR
	@mkdir -p bin
endef

# Build for Windows (x86_64)
windows-x86_64:
	@echo "Building for Windows with GOOS=windows, GOARCH=amd64..."
	$(call MAKE_BIN_DIR)
	@echo "Using zig cross compiler for Windows"
	@echo "Building with CC="zig cc -target x86_64-windows""
	CGO_ENABLED=1 GOOS=windows GOARCH=amd64 CC="zig cc -target x86_64-windows" \
		go build -o bin/$(BINARY_BASE)_windows-x86_64.exe .

# Build for Windows (ARM)
windows-arm:
	@echo "Building for Windows (ARM) with GOOS=windows, GOARCH=arm64..."
	$(call MAKE_BIN_DIR)
	@echo "Using zig cross compiler for Windows"
	@echo "Building with CC="zig cc -target aarch64-windows""
	CGO_ENABLED=1 GOOS=windows GOARCH=arm64 CC="zig cc -target aarch64-windows" \
		go build -o bin/$(BINARY_BASE)-windows-arm.exe .

# Build for Linux (x86_64)
linux-x86_64:
	@echo "Building for Linux with GOOS=linux, GOARCH=amd64..."
	$(call MAKE_BIN_DIR)
	@echo "Using gcc for Linux"
	@echo "Building with CC="zig cc -target x86_64-linux""
	CGO_ENABLED=1 GOOS=linux GOARCH=amd64 CC="zig cc -target x86_64-linux" \
		go build -o bin/$(BINARY_BASE)-linux-x86_64 .

# Build for Linux (ARM)
linux-arm:
	@echo "Building for Linux (ARM) with GOOS=linux, GOARCH=arm64..."
	$(call MAKE_BIN_DIR)
	@echo "Using gcc for Linux (ARM)"
	@echo "Building with CC="zig cc -target aarch64-linux""
	CGO_ENABLED=1 GOOS=linux GOARCH=arm64 CC="zig cc -target aarch64-linux" \
		go build -o bin/$(BINARY_BASE)-linux-arm .

# Build for macOS (ARM)
macos-arm:
	@echo "Building for macOS (ARM) with GOOS=darwin, GOARCH=arm64..."
	$(call MAKE_BIN_DIR)
	@echo "Using clang for macOS (ARM)"
	@echo "Building with CC=clang"
	CGO_ENABLED=1 GOOS=darwin GOARCH=arm64 CC=clang \
		go build -o bin/$(BINARY_BASE)-macos-arm .

# Build all targets
all: windows-x86_64 windows-arm linux-x86_64 linux-arm macos-x86_64 macos-arm

# Clean target to remove generated binaries and bin folder if needed
clean:
	@echo "Cleaning generated binaries..."
	@rm -rf bin

.PHONY: windows-x86_64 windows-arm linux-x86_64 linux-arm macos-arm all clean