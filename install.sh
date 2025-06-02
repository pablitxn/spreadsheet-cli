#!/bin/bash
# SpreadsheetCLI Installation Script

echo "=== SpreadsheetCLI Installation ==="
echo

# Check if .NET is installed
if ! command -v dotnet &> /dev/null; then
    echo "Error: .NET SDK is not installed."
    echo "Please install .NET 9.0 SDK from: https://dotnet.microsoft.com/download"
    exit 1
fi

# Build the project
echo "Building SpreadsheetCLI..."
dotnet build -c Release

if [ $? -ne 0 ]; then
    echo "Error: Build failed."
    exit 1
fi

# Create symlink for global access
INSTALL_DIR="/usr/local/bin"
EXECUTABLE_PATH="$(pwd)/bin/Release/net9.0/linux-x64/ssllm"

if [ -f "$EXECUTABLE_PATH" ]; then
    echo
    echo "Creating symlink..."
    sudo ln -sf "$EXECUTABLE_PATH" "$INSTALL_DIR/ssllm"
    
    if [ $? -eq 0 ]; then
        echo
        echo "✓ SpreadsheetCLI installed successfully!"
        echo
        echo "You can now use 'ssllm' from anywhere in your terminal."
        echo
        echo "Examples:"
        echo "  ssllm --help                     # Show help"
        echo "  ssllm interactive                # Start interactive mode"
        echo "  ssllm query <file> <query>       # Run a single query"
        echo "  ssllm browse                     # Browse and select files"
        echo "  ssllm test --auto                # Run ground truth tests"
        echo
    else
        echo "Error: Failed to create symlink. You may need to run this script with sudo."
        exit 1
    fi
else
    echo "Error: Executable not found at $EXECUTABLE_PATH"
    exit 1
fi