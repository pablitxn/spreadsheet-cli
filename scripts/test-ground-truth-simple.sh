#!/bin/bash

# Simplified script to test spreadsheet CLI against ground truth data
# Uses only bash and basic tools, no Python dependencies

# Configuration
DATASET_FILE="./scripts/dataset/expanded_dataset_moved.xlsx"
PROJECT_DIR="."
REPORT_FILE="./scripts/dataset/accuracy_report_$(date +%Y%m%d_%H%M%S).txt"
AUDIT_LOG_DIR="./logs/audit"

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

# Manual test cases from ground truth (first few for testing)
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
    echo ""
    echo "=== TEST RESULTS ==="
    echo ""
} > "$REPORT_FILE"

# Function to extract answer from audit log
extract_answer_from_log() {
    local log_file="$1"
    if [ -f "$log_file" ]; then
        # Extract the answer from the audit log JSON
        grep -o '"Answer":"[^"]*"' "$log_file" | cut -d'"' -f4
    fi
}

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

# Create marker for timestamp
touch /tmp/test_start_marker_$$

# Process each test case
print_header "Running tests..."
echo ""

for i in "${!TEST_QUESTIONS[@]}"; do
    QUESTION="${TEST_QUESTIONS[$i]}"
    EXPECTED_ANSWER="${EXPECTED_ANSWERS[$i]}"
    TEST_NUM=$((i + 1))
    
    print_info "Test $TEST_NUM/$TOTAL_TESTS: $QUESTION"
    
    # Run the query
    RESULT=$("$EXECUTABLE" "$DATASET_FILE" "$QUESTION" 2>&1)
    EXIT_CODE=$?
    
    if [ $EXIT_CODE -ne 0 ]; then
        print_error "Query execution failed!"
        echo "Test $TEST_NUM: FAILED (Execution Error)" >> "$REPORT_FILE"
        echo "  Question: $QUESTION" >> "$REPORT_FILE"
        echo "  Expected: $EXPECTED_ANSWER" >> "$REPORT_FILE"
        echo "  Error: $RESULT" >> "$REPORT_FILE"
        echo "" >> "$REPORT_FILE"
        ((FAILED++))
        continue
    fi
    
    # Wait a moment for audit log to be written
    sleep 1
    
    # Extract answer directly from the command output
    ACTUAL_ANSWER=$(echo "$RESULT" | grep -o '"Answer":"[^"]*"' | cut -d'"' -f4)
    
    if [ -n "$ACTUAL_ANSWER" ]; then
        # Compare answers
        if compare_answers "$ACTUAL_ANSWER" "$EXPECTED_ANSWER"; then
            print_success "PASSED - Expected: $EXPECTED_ANSWER, Got: $ACTUAL_ANSWER"
            echo "Test $TEST_NUM: PASSED" >> "$REPORT_FILE"
            echo "  Question: $QUESTION" >> "$REPORT_FILE"
            echo "  Expected: $EXPECTED_ANSWER" >> "$REPORT_FILE"
            echo "  Actual: $ACTUAL_ANSWER" >> "$REPORT_FILE"
            echo "" >> "$REPORT_FILE"
            ((PASSED++))
        else
            print_error "FAILED - Expected: $EXPECTED_ANSWER, Got: $ACTUAL_ANSWER"
            echo "Test $TEST_NUM: FAILED" >> "$REPORT_FILE"
            echo "  Question: $QUESTION" >> "$REPORT_FILE"
            echo "  Expected: $EXPECTED_ANSWER" >> "$REPORT_FILE"
            echo "  Actual: $ACTUAL_ANSWER" >> "$REPORT_FILE"
            echo "" >> "$REPORT_FILE"
            ((FAILED++))
        fi
    else
        print_warning "No answer found in output"
        echo "Test $TEST_NUM: FAILED (No Answer Found)" >> "$REPORT_FILE"
        echo "  Question: $QUESTION" >> "$REPORT_FILE"
        echo "  Expected: $EXPECTED_ANSWER" >> "$REPORT_FILE"
        echo "" >> "$REPORT_FILE"
        ((FAILED++))
    fi
    
    echo ""
done

# Clean up
rm -f /tmp/test_start_marker_$$

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
    echo "Report saved to: $REPORT_FILE"
} | tee -a "$REPORT_FILE"

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
print_info "Detailed report saved to: $REPORT_FILE"

# Show first few lines of report
echo ""
print_header "=== SAMPLE FROM REPORT ==="
head -n 30 "$REPORT_FILE"