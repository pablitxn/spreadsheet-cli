#!/bin/bash

# Script to test spreadsheet CLI against ground truth data with organized logging
# Creates a dedicated folder for each test run with all logs

# Configuration
DATASET_FILE="./scripts/dataset/expanded_dataset_moved.xlsx"
PROJECT_DIR="."
LOGS_BASE_DIR="./logs"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
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

print_header() {
    echo -e "${CYAN}$1${NC}"
}

# Check if required files exist
if [ ! -f "$DATASET_FILE" ]; then
    print_error "Dataset file '$DATASET_FILE' not found!"
    exit 1
fi

# Create logs directory if it doesn't exist
mkdir -p "$LOGS_BASE_DIR"

# Extract dataset filename without path and extension
DATASET_NAME=$(basename "$DATASET_FILE" .xlsx)

# Create test run directory with timestamp and dataset name
TIMESTAMP=$(date +%Y%m%d_%H%M%S)
TEST_RUN_DIR="$LOGS_BASE_DIR/test_${DATASET_NAME}_${TIMESTAMP}"
mkdir -p "$TEST_RUN_DIR"

# Create subdirectories
mkdir -p "$TEST_RUN_DIR/audit"
mkdir -p "$TEST_RUN_DIR/debug"
mkdir -p "$TEST_RUN_DIR/query_outputs"

# Set report file path
REPORT_FILE="$TEST_RUN_DIR/accuracy_report.txt"

print_info "Test run directory: $TEST_RUN_DIR"

# Create test info file
{
    echo "Test Run Information"
    echo "===================="
    echo "Date: $(date)"
    echo "Dataset: $DATASET_FILE"
    echo "Test Directory: $TEST_RUN_DIR"
    echo ""
} > "$TEST_RUN_DIR/test_info.txt"

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

# Manual test cases from ground truth
declare -a TEST_QUESTIONS=(
    "What is the total sum of TotalBaseIncome?"
    "What is the average TotalBaseIncome?"
    "What is the maximum value of TotalBaseIncome?"
    "What is the average Quantity?"
    "What is the total sum of Quantity?"
    "What percentage of records have PaymentType = 'CASH'?"
    "Which SecurityID shows the greatest total TotalBaseIncome?"
    "What is the total number of unique SecurityIDs?"
    "Which SecurityID has the highest average TotalBaseIncome?"
    "What is the total TotalBaseIncome for PaymentType = 'CASH'?"
)

declare -a EXPECTED_ANSWERS=(
    "12977.52"
    "91.13"
    "3700.00"
    "4988.27"
    "710248"
    "48.69"
    "MSFT"
    "5"
    "MSFT"
    "5869.48"
)

# Initialize counters
TOTAL_TESTS=${#TEST_QUESTIONS[@]}
PASSED=0
FAILED=0

# Start report
{
    echo "=== GROUND TRUTH ACCURACY REPORT ==="
    echo "Date: $(date)"
    echo "Dataset: $DATASET_FILE"
    echo "Total Test Cases: $TOTAL_TESTS"
    echo "Test Directory: $TEST_RUN_DIR"
    echo ""
    echo "=== TEST RESULTS ==="
    echo ""
} > "$REPORT_FILE"

# Function to normalize numeric values for comparison
normalize_numeric() {
    local value="$1"
    # Remove commas, dollar signs, percent signs, and extra spaces
    echo "$value" | sed 's/[$,%]//g' | sed 's/^[[:space:]]*//;s/[[:space:]]*$//'
}

# Function to compare answers
compare_answers() {
    local actual="$1"
    local expected="$2"
    
    # Normalize both values
    actual=$(normalize_numeric "$actual")
    expected=$(normalize_numeric "$expected")
    
    # Try exact match first
    if [ "$actual" = "$expected" ]; then
        return 0
    fi
    
    # Try case-insensitive match for text
    if [ "${actual,,}" = "${expected,,}" ]; then
        return 0
    fi
    
    # Try numeric comparison with tolerance (without bc)
    if [[ "$actual" =~ ^[0-9]+(\.[0-9]+)?$ ]] && [[ "$expected" =~ ^[0-9]+(\.[0-9]+)?$ ]]; then
        # Compare integers (remove decimal part for simple comparison)
        local actual_int=${actual%.*}
        local expected_int=${expected%.*}
        local diff=$((actual_int > expected_int ? actual_int - expected_int : expected_int - actual_int))
        # Allow difference of 1 for rounding
        if [ $diff -le 1 ]; then
            return 0
        fi
    fi
    
    return 1
}

# Create environment variables for the app to use our test directory
export AUDIT_LOG_DIR="$TEST_RUN_DIR/audit"
export DEBUG_LOG_DIR="$TEST_RUN_DIR/debug"

# Process each test case
print_header "Running tests..."
echo ""

for i in "${!TEST_QUESTIONS[@]}"; do
    QUESTION="${TEST_QUESTIONS[$i]}"
    EXPECTED_ANSWER="${EXPECTED_ANSWERS[$i]}"
    TEST_NUM=$((i + 1))
    
    print_info "Test $TEST_NUM/$TOTAL_TESTS: $QUESTION"
    
    # Create query-specific file names
    QUERY_OUTPUT_FILE="$TEST_RUN_DIR/query_outputs/test_${TEST_NUM}_output.json"
    QUERY_ERROR_FILE="$TEST_RUN_DIR/query_outputs/test_${TEST_NUM}_error.txt"
    
    # Run the query and capture all output
    RESULT=$("$EXECUTABLE" "$DATASET_FILE" "$QUESTION" 2>"$QUERY_ERROR_FILE")
    EXIT_CODE=$?
    
    # Save the output
    echo "$RESULT" > "$QUERY_OUTPUT_FILE"
    
    if [ $EXIT_CODE -ne 0 ]; then
        print_error "Query execution failed!"
        echo "Test $TEST_NUM: FAILED (Execution Error)" >> "$REPORT_FILE"
        echo "  Question: $QUESTION" >> "$REPORT_FILE"
        echo "  Expected: $EXPECTED_ANSWER" >> "$REPORT_FILE"
        echo "  Error: See $QUERY_ERROR_FILE" >> "$REPORT_FILE"
        echo "" >> "$REPORT_FILE"
        ((FAILED++))
        continue
    fi
    
    # Extract answer from the command output
    ACTUAL_ANSWER=$(echo "$RESULT" | grep -o '"Answer":"[^"]*"' | cut -d'"' -f4)
    
    # Also save the full JSON output with test info
    {
        echo "{"
        echo "  \"test_number\": $TEST_NUM,"
        echo "  \"question\": \"$QUESTION\","
        echo "  \"expected_answer\": \"$EXPECTED_ANSWER\","
        echo "  \"actual_answer\": \"$ACTUAL_ANSWER\","
        echo "  \"result\":"
        echo "$RESULT"
        echo "}"
    } > "$TEST_RUN_DIR/query_outputs/test_${TEST_NUM}_full.json"
    
    if [ -n "$ACTUAL_ANSWER" ]; then
        # Compare answers
        if compare_answers "$ACTUAL_ANSWER" "$EXPECTED_ANSWER"; then
            print_success "PASSED - Expected: $EXPECTED_ANSWER, Got: $ACTUAL_ANSWER"
            echo "Test $TEST_NUM: PASSED" >> "$REPORT_FILE"
            echo "  Question: $QUESTION" >> "$REPORT_FILE"
            echo "  Expected: $EXPECTED_ANSWER" >> "$REPORT_FILE"
            echo "  Actual: $ACTUAL_ANSWER" >> "$REPORT_FILE"
            echo "  Output: $QUERY_OUTPUT_FILE" >> "$REPORT_FILE"
            echo "" >> "$REPORT_FILE"
            ((PASSED++))
        else
            print_error "FAILED - Expected: $EXPECTED_ANSWER, Got: $ACTUAL_ANSWER"
            echo "Test $TEST_NUM: FAILED" >> "$REPORT_FILE"
            echo "  Question: $QUESTION" >> "$REPORT_FILE"
            echo "  Expected: $EXPECTED_ANSWER" >> "$REPORT_FILE"
            echo "  Actual: $ACTUAL_ANSWER" >> "$REPORT_FILE"
            echo "  Output: $QUERY_OUTPUT_FILE" >> "$REPORT_FILE"
            echo "" >> "$REPORT_FILE"
            ((FAILED++))
        fi
    else
        print_warning "No answer found in output"
        echo "Test $TEST_NUM: FAILED (No Answer Found)" >> "$REPORT_FILE"
        echo "  Question: $QUESTION" >> "$REPORT_FILE"
        echo "  Expected: $EXPECTED_ANSWER" >> "$REPORT_FILE"
        echo "  Output: $QUERY_OUTPUT_FILE" >> "$REPORT_FILE"
        echo "" >> "$REPORT_FILE"
        ((FAILED++))
    fi
    
    # Copy any new logs generated by this query
    # Note: In a real implementation, the app should write directly to our test directory
    # For now, we'll check if any new logs were created in the default location
    if [ -d "./logs/audit" ]; then
        find "./logs/audit" -name "*.json" -newer "$QUERY_OUTPUT_FILE" -exec cp {} "$TEST_RUN_DIR/audit/" \; 2>/dev/null
    fi
    if [ -d "./logs" ]; then
        find "./logs" -name "debug_*.log" -newer "$QUERY_OUTPUT_FILE" -exec cp {} "$TEST_RUN_DIR/debug/" \; 2>/dev/null
    fi
    
    echo ""
done

# Calculate accuracy without bc
if [ $TOTAL_TESTS -gt 0 ]; then
    ACCURACY=$((PASSED * 100 / TOTAL_TESTS))
else
    ACCURACY=0
fi

# Append summary to report
{
    echo ""
    echo "=== SUMMARY ==="
    echo "Total Tests: $TOTAL_TESTS"
    echo "Passed: $PASSED"
    echo "Failed: $FAILED"
    echo "Accuracy: ${ACCURACY}%"
    echo ""
    echo "Test Directory: $TEST_RUN_DIR"
} | tee -a "$REPORT_FILE"

# Create summary file
{
    echo "Test Summary"
    echo "============"
    echo "Total Tests: $TOTAL_TESTS"
    echo "Passed: $PASSED"
    echo "Failed: $FAILED"
    echo "Accuracy: ${ACCURACY}%"
    echo ""
    echo "Directory Structure:"
    echo "  accuracy_report.txt - Full test report"
    echo "  test_info.txt - Test run information"
    echo "  audit/ - Audit logs for each query"
    echo "  debug/ - Debug logs"
    echo "  query_outputs/ - Raw outputs from each query"
} > "$TEST_RUN_DIR/summary.txt"

# Display summary
echo ""
print_header "=== TEST SUMMARY ==="
print_info "Total Tests: $TOTAL_TESTS"
print_success "Passed: $PASSED"
print_error "Failed: $FAILED"
echo ""
if [ $ACCURACY -ge 70 ]; then
    print_success "Accuracy: ${ACCURACY}%"
elif [ $ACCURACY -ge 50 ]; then
    print_warning "Accuracy: ${ACCURACY}%"
else
    print_error "Accuracy: ${ACCURACY}%"
fi
echo ""
print_info "All logs saved to: $TEST_RUN_DIR"
print_info "View report: cat $REPORT_FILE"
print_info "View summary: cat $TEST_RUN_DIR/summary.txt"

# Show tree structure of the test directory if tree command is available
if command -v tree &> /dev/null; then
    echo ""
    print_header "=== TEST DIRECTORY STRUCTURE ==="
    tree "$TEST_RUN_DIR" -L 2
fi