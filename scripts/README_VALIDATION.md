# LLM-Based Test Validation

This directory contains tools for validating test results using LLM (Large Language Model) to intelligently extract and compare answers from the CLI output.

## Overview

The test validation system uses OpenAI's GPT-4o-mini model with structured outputs to:
- Extract answers from various fields in the JSON output (Answer, MachineAnswer, Reasoning, etc.)
- Handle different answer formats (numbers with commas, percentages, entity names)
- Provide confidence scores and explanations for validation decisions
- Debug where the answer was found in the output

## Components

### 1. `validate_test_result.py`
Python script that uses OpenAI API to validate test results.

**Usage:**
```bash
python3 validate_test_result.py "<question>" "<expected_answer>" "<actual_output>"
```

**Requirements:**
- Python 3.6+
- OpenAI Python library: `pip install openai`
- Environment variable: `OPENAI_API_KEY`

**Output:**
JSON object with:
- `is_correct`: boolean indicating if the test passed
- `extracted_answer`: the answer found in the output
- `explanation`: why the test passed or failed
- `confidence`: confidence score (0-1)
- `answer_location`: where the answer was found

### 2. `TestResultValidationService.cs`
C# service implementing the same validation logic for use within the application.

### 3. Integration with `test-ground-truth.sh`
The test script automatically uses LLM validation when:
- The Python script exists
- Python 3 is available
- OpenAI library is installed
- `OPENAI_API_KEY` is set

If LLM validation is not available, it falls back to pattern matching.

## Setup

1. Install Python dependencies:
   ```bash
   pip install openai
   ```

2. Set your OpenAI API key:
   ```bash
   export OPENAI_API_KEY="your-api-key-here"
   ```

3. Run the test script:
   ```bash
   ./scripts/test-ground-truth.sh
   ```

## Benefits

- **Better Answer Extraction**: The LLM can understand context and extract answers even when they're embedded in explanations
- **Flexible Matching**: Handles different number formats, units, and text variations
- **Debugging Support**: Shows where the answer was found and why validation passed/failed
- **Confidence Scores**: Provides confidence levels for each validation

## Validation Rules

1. **Numeric Comparison**:
   - Allows small rounding differences (±0.01)
   - Ignores formatting (commas, dollar signs, percent signs)
   - "12,977.52" matches "12977.52" or "$12977.52"

2. **Text Comparison**:
   - Case-insensitive matching
   - "MSFT" matches "msft" or "Msft"

3. **Answer Priority**:
   - First checks MachineAnswer field
   - Then Answer and SimpleAnswer fields
   - Finally searches in Reasoning field

4. **Special Cases**:
   - Percentage questions: accepts with or without % sign
   - Count questions: looks for "X rows", "X records"
   - Which X has Y questions: accepts either X or Y as answer