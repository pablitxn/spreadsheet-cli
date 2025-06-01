#!/bin/bash

# Script to generate ground truth values from actual dataset
# This ensures test values always match the current dataset

# Configuration
DATASET_FILE="./scripts/dataset/expanded_dataset_moved.xlsx"
OUTPUT_FILE="./scripts/dataset/ground_truth_values.json"
EXECUTABLE=""

# Determine the executable path
if [ -f "./bin/Release/net9.0/linux-x64/ssllm" ]; then
    EXECUTABLE="./bin/Release/net9.0/linux-x64/ssllm"
elif [ -f "./bin/Debug/net9.0/linux-x64/ssllm" ]; then
    EXECUTABLE="./bin/Debug/net9.0/linux-x64/ssllm"
elif [ -f "./bin/Release/net9.0/ssllm" ]; then
    EXECUTABLE="./bin/Release/net9.0/ssllm"
elif [ -f "./bin/Debug/net9.0/ssllm" ]; then
    EXECUTABLE="./bin/Debug/net9.0/ssllm"
else
    echo "Error: Executable not found! Please build the project first."
    exit 1
fi

# Check if dataset exists
if [ ! -f "$DATASET_FILE" ]; then
    echo "Error: Dataset file '$DATASET_FILE' not found!"
    exit 1
fi

# Calculate dataset hash
DATASET_HASH=$(sha256sum "$DATASET_FILE" 2>/dev/null | cut -d' ' -f1)
if [ -z "$DATASET_HASH" ]; then
    # Fallback for macOS
    DATASET_HASH=$(shasum -a 256 "$DATASET_FILE" 2>/dev/null | cut -d' ' -f1)
fi

echo "Generating ground truth values for dataset: $DATASET_FILE"
echo "Dataset SHA-256: $DATASET_HASH"

# Test questions
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

# Start JSON output
cat > "$OUTPUT_FILE" << EOF
{
  "dataset_hash": "$DATASET_HASH",
  "dataset_file": "$DATASET_FILE",
  "generated_date": "$(date -u +%Y-%m-%dT%H:%M:%SZ)",
  "test_cases": [
EOF

# Process each question
FIRST=true
for i in "${!TEST_QUESTIONS[@]}"; do
    QUESTION="${TEST_QUESTIONS[$i]}"
    echo "Processing question $((i+1))/${#TEST_QUESTIONS[@]}: $QUESTION"
    
    # Run the query
    RESULT=$("$EXECUTABLE" "$DATASET_FILE" "$QUESTION" 2>&1)
    
    if [ $? -eq 0 ]; then
        # Extract answer from JSON output
        ANSWER=$(echo "$RESULT" | grep -o '"Answer":"[^"]*"' | cut -d'"' -f4)
        
        if [ -n "$ANSWER" ]; then
            # Add comma if not first item
            if [ "$FIRST" = false ]; then
                echo "," >> "$OUTPUT_FILE"
            fi
            FIRST=false
            
            # Write test case
            cat >> "$OUTPUT_FILE" << EOF
    {
      "question": "$QUESTION",
      "expected_answer": "$ANSWER"
    }
EOF
        else
            echo "Warning: No answer found for question: $QUESTION"
        fi
    else
        echo "Error: Failed to execute query: $QUESTION"
    fi
done

# Close JSON
cat >> "$OUTPUT_FILE" << EOF

  ]
}
EOF

echo ""
echo "Ground truth values generated successfully!"
echo "Output file: $OUTPUT_FILE"
echo ""
echo "You can now use these values in your test scripts."
echo "Remember to regenerate this file whenever the dataset changes."