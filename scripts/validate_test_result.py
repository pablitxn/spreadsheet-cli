#!/usr/bin/env python3
"""
Test result validation using OpenAI API with structured outputs.
This script validates if the actual output from the CLI correctly answers the given question.
"""

import os
import sys
import json
import argparse
import logging
from typing import Dict, Any, Optional
from dataclasses import dataclass, asdict

# Try to import OpenAI library
try:
    from openai import OpenAI
except ImportError:
    print("Error: OpenAI library not installed. Install with: pip install openai", file=sys.stderr)
    sys.exit(1)

# Configure logging
logging.basicConfig(level=logging.WARNING, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)


@dataclass
class TestValidationResult:
    """Result of test validation using LLM"""
    is_correct: bool
    extracted_answer: str
    explanation: str
    confidence: float
    answer_location: str


class TestValidator:
    """Validates test results using OpenAI API"""
    
    def __init__(self, api_key: Optional[str] = None):
        """Initialize validator with API key"""
        self.api_key = api_key or os.environ.get('OPENAI_API_KEY')
        if not self.api_key:
            raise ValueError("OpenAI API key not found. Set OPENAI_API_KEY environment variable.")
        
        self.client = OpenAI(api_key=self.api_key)
    
    def validate_test_result(self, question: str, expected_answer: str, actual_output: str) -> TestValidationResult:
        """Validate if actual output correctly answers the question"""
        
        prompt = f"""You are a test validator for a spreadsheet analysis CLI. Your job is to determine if the actual output correctly answers the given question.

Question: {question}
Expected Answer: {expected_answer}

The CLI output may contain the answer in various fields:
- Answer: Direct answer field
- MachineAnswer: Machine-readable answer (numeric values, entity names)
- Reasoning: Explanation that may contain the answer
- SimpleAnswer: Alternative answer field
- HumanExplanation: Human-readable explanation

Actual CLI Output:
{actual_output}

VALIDATION RULES:
1. Numeric Comparison:
   - Allow for small rounding differences (±0.01 for decimals)
   - Ignore formatting differences (commas, dollar signs, percent signs)
   - "12977.52" matches "12,977.52" or "$12977.52"
   - For percentages: "48.69" matches "48.69%" 
   
2. Text Comparison:
   - Case-insensitive matching
   - "MSFT" matches "msft" or "Msft"
   
3. Answer Extraction Priority:
   - First check MachineAnswer field (most reliable for numeric/entity answers)
   - Then check Answer and SimpleAnswer fields
   - Finally search in Reasoning field for answer patterns
   - Look for patterns like "The answer is X", "total is X", "average is X"
   
4. Special Cases:
   - For "Which X has highest/lowest Y" questions: Accept either just the X value or the Y value
   - For percentage questions: Accept with or without % sign
   - For count questions: Look for "X rows", "X records", "X unique values"
   
IMPORTANT: Be flexible in extraction but strict in validation. The answer must be semantically correct.

Extract the actual answer from the output and determine if it matches the expected answer."""

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "user", "content": prompt}
                ],
                response_format={
                    "type": "json_schema",
                    "json_schema": {
                        "name": "test_validation",
                        "strict": True,
                        "schema": {
                            "type": "object",
                            "properties": {
                                "is_correct": {
                                    "type": "boolean",
                                    "description": "Whether the actual output correctly answers the question"
                                },
                                "extracted_answer": {
                                    "type": "string",
                                    "description": "The answer extracted from the actual output"
                                },
                                "explanation": {
                                    "type": "string",
                                    "description": "Brief explanation of why the test passed or failed"
                                },
                                "confidence": {
                                    "type": "number",
                                    "description": "Confidence level of the validation (0-1)",
                                    "minimum": 0,
                                    "maximum": 1
                                },
                                "answer_location": {
                                    "type": "string",
                                    "description": "Where the answer was found (e.g., 'MachineAnswer field', 'Reasoning field')"
                                }
                            },
                            "required": ["is_correct", "extracted_answer", "explanation", "confidence", "answer_location"],
                            "additionalProperties": False
                        }
                    }
                },
                temperature=0.1
            )
            
            result_data = json.loads(response.choices[0].message.content)
            return TestValidationResult(**result_data)
            
        except Exception as e:
            logger.error(f"Error during validation: {e}")
            return TestValidationResult(
                is_correct=False,
                extracted_answer="Error",
                explanation=f"Validation failed due to error: {str(e)}",
                confidence=0.0,
                answer_location="N/A"
            )


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(description='Validate test results using OpenAI API')
    parser.add_argument('question', help='The question that was asked')
    parser.add_argument('expected_answer', help='The expected answer from ground truth')
    parser.add_argument('actual_output', help='The actual output from the CLI')
    parser.add_argument('--api-key', help='OpenAI API key (optional, can use OPENAI_API_KEY env var)')
    parser.add_argument('--verbose', '-v', action='store_true', help='Enable verbose output')
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.INFO)
    
    try:
        validator = TestValidator(api_key=args.api_key)
        result = validator.validate_test_result(
            question=args.question,
            expected_answer=args.expected_answer,
            actual_output=args.actual_output
        )
        
        # Output result as JSON
        print(json.dumps(asdict(result)))
        
        # Exit with appropriate code
        sys.exit(0 if result.is_correct else 1)
        
    except Exception as e:
        logger.error(f"Fatal error: {e}")
        error_result = {
            "is_correct": False,
            "extracted_answer": "Error",
            "explanation": str(e),
            "confidence": 0.0,
            "answer_location": "N/A"
        }
        print(json.dumps(error_result))
        sys.exit(1)


if __name__ == '__main__':
    main()