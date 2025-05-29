#!/bin/bash

# Script to quickly query the spreadsheet CLI
# Usage: ./query.sh "your question here"
# Example: ./query.sh "Which SecurityID shows the greatest standard deviation of TotalBaseIncome?"

# Configuration - modify these paths as needed
DATASET_FILE="expanded_dataset_moved.xlsx"
PROJECT_DIR="."

# Colors for better output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Function to print colored output
print_info() {
    echo -e "${BLUE}ℹ${NC} $1"
}

print_success() {
    echo -e "${GREEN}✓${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}⚠${NC} $1"
}

print_error() {
    echo -e "${RED}✗${NC} $1"
}

# Check if question is provided
if [ -z "$1" ]; then
    print_error "No question provided!"
    echo
    echo "Usage: $0 \"your question here\""
    echo
    echo "Examples:"
    echo "  $0 \"Which SecurityID shows the greatest standard deviation of TotalBaseIncome?\""
    echo "  $0 \"What percentage of items have quantity > 1000?\""
    echo "  $0 \"Show me the top 5 records by some criteria\""
    exit 1
fi

# Check if dataset file exists
if [ ! -f "$DATASET_FILE" ]; then
    print_error "Dataset file '$DATASET_FILE' not found!"
    print_info "Please ensure the file exists or modify the DATASET_FILE variable in this script."
    exit 1
fi

# Change to project directory if specified
if [ "$PROJECT_DIR" != "." ]; then
    cd "$PROJECT_DIR" || {
        print_error "Failed to change to project directory: $PROJECT_DIR"
        exit 1
    }
fi

# Determine the executable path
EXECUTABLE=""
if [ -f "./bin/Release/net9.0/linux-x64/ssllm" ]; then
    EXECUTABLE="./bin/Release/net9.0/linux-x64/ssllm"
elif [ -f "./bin/Debug/net9.0/linux-x64/ssllm" ]; then
    EXECUTABLE="./bin/Debug/net9.0/linux-x64/ssllm"
elif [ -f "./bin/Release/net9.0/ssllm" ]; then
    EXECUTABLE="./bin/Release/net9.0/ssllm"
elif [ -f "./bin/Debug/net9.0/ssllm" ]; then
    EXECUTABLE="./bin/Debug/net9.0/ssllm"
else
    print_error "Executable not found! Please build the project first."
    exit 1
fi

# Display query information
echo
print_info "Dataset: $DATASET_FILE"
print_info "Query: $1"
echo
print_success "🚀 Executing query..."
echo

# Run the query with proper error handling
RESULT=$("$EXECUTABLE" "$DATASET_FILE" "$1" 2>&1)
EXIT_CODE=$?

if [ $EXIT_CODE -ne 0 ]; then
    print_error "Query execution failed!"
    echo "$RESULT"
    exit 1
fi

# Extract only the JSON output
# Look for content between first { and last }
JSON_OUTPUT=$(echo "$RESULT" | awk '/^\{/{p=1} p{print} /^\}/{p=0}')

if [ -z "$JSON_OUTPUT" ]; then
    # If no JSON found, print all output
    echo "$RESULT"
else
    # Pretty print the JSON
    echo "$JSON_OUTPUT" | python3 -m json.tool 2>/dev/null || echo "$JSON_OUTPUT"
fi

print_success "Query completed successful" 