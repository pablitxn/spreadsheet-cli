#!/bin/bash

# Script to test spreadsheet CLI against ground truth data
# Compares actual results with expected answers from ground_truth_expanded_dataset_moved.xlsx

# Configuration
DATASET_FILE="./scripts/dataset/expanded_dataset_moved.xlsx"
GROUND_TRUTH_FILE="./scripts/dataset/ground_truth_expanded_dataset_moved.xlsx"
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

if [ ! -f "$GROUND_TRUTH_FILE" ]; then
    print_error "Ground truth file '$GROUND_TRUTH_FILE' not found!"
    exit 1
fi

# Check if Python is available for parsing Excel
if ! command -v python3 &> /dev/null; then
    print_error "Python3 is required to parse Excel files"
    exit 1
fi

# Check if openpyxl is available
python3 -c "import openpyxl" &> /dev/null || {
    print_error "Python package 'openpyxl' is required but not installed"
    print_info "Please install it manually: python3 -m pip install openpyxl"
    print_info "Or use your system package manager"
    exit 1
}

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

# Create Python script to extract questions and answers
cat > /tmp/extract_ground_truth.py << 'EOF'
import openpyxl
import json
import sys

def extract_ground_truth(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active
    
    test_cases = []
    
    # Skip header row
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:  # Question and Answer columns
            test_cases.append({
                'question': str(row[0]).strip(),
                'expected_answer': str(row[1]).strip(),
                'notes': str(row[2]).strip() if row[2] else ''
            })
    
    return test_cases

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python extract_ground_truth.py <ground_truth_file>")
        sys.exit(1)
    
    test_cases = extract_ground_truth(sys.argv[1])
    print(json.dumps(test_cases))
EOF

# Extract test cases from ground truth
print_info "Extracting test cases from ground truth..."
TEST_CASES=$(python3 /tmp/extract_ground_truth.py "$GROUND_TRUTH_FILE")

if [ -z "$TEST_CASES" ]; then
    print_error "No test cases found in ground truth file!"
    exit 1
fi

# Count total test cases
TOTAL_TESTS=$(echo "$TEST_CASES" | python3 -c "import json, sys; print(len(json.load(sys.stdin)))")
print_success "Found $TOTAL_TESTS test cases"

# Initialize counters
PASSED=0
FAILED=0
ERRORS=0

# Start report
{
    echo "=== GROUND TRUTH ACCURACY REPORT ==="
    echo "Date: $(date)"
    echo "Dataset: $DATASET_FILE"
    echo "Ground Truth: $GROUND_TRUTH_FILE"
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
        local answer=$(grep -o '"Answer":"[^"]*"' "$log_file" | cut -d'"' -f4)
        echo "$answer"
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
    
    # Try numeric comparison with tolerance
    if [[ "$actual" =~ ^[0-9]+(\.[0-9]+)?$ ]] && [[ "$expected" =~ ^[0-9]+(\.[0-9]+)?$ ]]; then
        # Compare as floating point with 0.01 tolerance
        local diff=$(echo "scale=6; a=$actual; e=$expected; if (a>e) a-e else e-a" | bc 2>/dev/null)
        if [ $? -eq 0 ] && [ "$(echo "$diff < 0.01" | bc)" -eq 1 ]; then
            return 0
        fi
    fi
    
    return 1
}

# Get latest audit log timestamp before running tests
BEFORE_TIMESTAMP=$(date +%s)

# Process each test case
print_header "Running tests..."
echo ""

echo "$TEST_CASES" | python3 -c "
import json
import sys

test_cases = json.load(sys.stdin)
for i, test in enumerate(test_cases):
    print(f'{i}|{test[\"question\"]}|{test[\"expected_answer\"]}|{test.get(\"notes\", \"\")}')
" | while IFS='|' read -r INDEX QUESTION EXPECTED_ANSWER NOTES; do
    
    TEST_NUM=$((INDEX + 1))
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
        ((ERRORS++))
        continue
    fi
    
    # Wait a moment for audit log to be written
    sleep 0.5
    
    # Find the latest audit log file
    LATEST_LOG=$(find "$AUDIT_LOG_DIR" -name "*.json" -newer /tmp/timestamp_marker_$$ 2>/dev/null | sort -r | head -1)
    
    if [ -z "$LATEST_LOG" ]; then
        # Try without timestamp filter
        LATEST_LOG=$(ls -t "$AUDIT_LOG_DIR"/*.json 2>/dev/null | head -1)
    fi
    
    if [ -n "$LATEST_LOG" ]; then
        ACTUAL_ANSWER=$(extract_answer_from_log "$LATEST_LOG")
        
        if [ -n "$ACTUAL_ANSWER" ]; then
            # Compare answers
            if compare_answers "$ACTUAL_ANSWER" "$EXPECTED_ANSWER"; then
                print_success "PASSED - Expected: $EXPECTED_ANSWER, Got: $ACTUAL_ANSWER"
                echo "Test $TEST_NUM: PASSED" >> "$REPORT_FILE"
                echo "  Question: $QUESTION" >> "$REPORT_FILE"
                echo "  Expected: $EXPECTED_ANSWER" >> "$REPORT_FILE"
                echo "  Actual: $ACTUAL_ANSWER" >> "$REPORT_FILE"
                if [ -n "$NOTES" ]; then
                    echo "  Notes: $NOTES" >> "$REPORT_FILE"
                fi
                echo "" >> "$REPORT_FILE"
                ((PASSED++))
            else
                print_error "FAILED - Expected: $EXPECTED_ANSWER, Got: $ACTUAL_ANSWER"
                echo "Test $TEST_NUM: FAILED" >> "$REPORT_FILE"
                echo "  Question: $QUESTION" >> "$REPORT_FILE"
                echo "  Expected: $EXPECTED_ANSWER" >> "$REPORT_FILE"
                echo "  Actual: $ACTUAL_ANSWER" >> "$REPORT_FILE"
                if [ -n "$NOTES" ]; then
                    echo "  Notes: $NOTES" >> "$REPORT_FILE"
                fi
                echo "" >> "$REPORT_FILE"
                ((FAILED++))
            fi
        else
            print_warning "No answer found in audit log"
            echo "Test $TEST_NUM: FAILED (No Answer Found)" >> "$REPORT_FILE"
            echo "  Question: $QUESTION" >> "$REPORT_FILE"
            echo "  Expected: $EXPECTED_ANSWER" >> "$REPORT_FILE"
            echo "  Log file: $LATEST_LOG" >> "$REPORT_FILE"
            echo "" >> "$REPORT_FILE"
            ((FAILED++))
        fi
    else
        print_warning "No audit log found for this query"
        echo "Test $TEST_NUM: FAILED (No Audit Log)" >> "$REPORT_FILE"
        echo "  Question: $QUESTION" >> "$REPORT_FILE"
        echo "  Expected: $EXPECTED_ANSWER" >> "$REPORT_FILE"
        echo "" >> "$REPORT_FILE"
        ((FAILED++))
    fi
    
    # Create timestamp marker for next test
    touch /tmp/timestamp_marker_$$
    
    echo ""
done

# Clean up
rm -f /tmp/extract_ground_truth.py
rm -f /tmp/timestamp_marker_$$

# Calculate accuracy
if [ $TOTAL_TESTS -gt 0 ]; then
    ACCURACY=$(echo "scale=2; $PASSED * 100 / $TOTAL_TESTS" | bc)
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
    echo "Errors: $ERRORS"
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
print_warning "Errors: $ERRORS"
echo ""
if [ $(echo "$ACCURACY >= 70" | bc) -eq 1 ]; then
    print_success "Accuracy: ${ACCURACY}%"
elif [ $(echo "$ACCURACY >= 50" | bc) -eq 1 ]; then
    print_warning "Accuracy: ${ACCURACY}%"
else
    print_error "Accuracy: ${ACCURACY}%"
fi
echo ""
print_info "Detailed report saved to: $REPORT_FILE"