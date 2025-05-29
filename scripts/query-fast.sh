#!/bin/bash

# Fast query script - skips build for quicker iterations
# Usage: ./query-fast.sh "your question here"

# Example queries for expanded_dataset_moved.xlsx:
# 1. What is the total Quantity for SecurityID 101121101 across the dataset?
# 2. What is the average TotalBaseIncome for rows with PaymentType 'DIV'?
# 3. What is the maximum TotalBaseIncome observed in the dataset?
# 4. Which SecurityID contributes the highest aggregate TotalBaseIncome?
# 5. How many unique PaymentType values are present?
# 6. What is the average Quantity for SIMON PROPERTY GROUP INC entries?
# 7. What is the total TotalBaseIncome for SecurityGroup 1?
# 8. What percentage of rows have Quantity greater than 1000?
# 9. What is the earliest ExDate in the dataset?
# 10. What is the TotalBaseIncome-to-Quantity ratio for SecurityID 828806109?
# 11. For which SecurityGroup is the variance of Quantity the highest, and what is that variance?
# 12. What is the difference between the median Quantity of DIV payments and the median Quantity of ACTUAL payments (DIV − ACTUAL)?
# 13. What is the ratio of the total TotalBaseIncome of ACTUAL rows to that of DIV rows?
# 14. What is the 90th percentile (P90) of Quantity across the dataset?
# 15. Which SecurityID shows the greatest standard deviation of TotalBaseIncome, and what is that standard deviation?
# 16. Which SecurityGroup derives the largest share of its TotalBaseIncome from DIV payments, and what is that share?
# 17. What is the average daily TotalBaseIncome over the full range of ExDates (earliest→latest, inclusive)?
# 18. How many rows record a negative TotalBaseIncome?
# 19. For SecurityID 101121101, what is the coefficient of variation of Quantity (std ÷ mean)?

if [ -z "$1" ]; then
    echo "Usage: ./query-fast.sh \"your question here\""
    echo "Example: ./query-fast.sh \"What is the total revenue?\""
    echo ""
    echo "Or use ./query-fast.sh examples to see example queries"
    exit 1
fi

# Check if user wants to see examples
if [ "$1" == "examples" ]; then
    echo "Example queries for expanded_dataset_moved.xlsx:"
    echo ""
    echo "1. \"What is the total Quantity for SecurityID 101121101 across the dataset?\""
    echo "2. \"What is the average TotalBaseIncome for rows with PaymentType 'DIV'?\""
    echo "3. \"What is the maximum TotalBaseIncome observed in the dataset?\""
    echo "4. \"Which SecurityID contributes the highest aggregate TotalBaseIncome?\""
    echo "5. \"How many unique PaymentType values are present?\""
    echo "6. \"What is the average Quantity for SIMON PROPERTY GROUP INC entries?\""
    echo "7. \"What is the total TotalBaseIncome for SecurityGroup 1?\""
    echo "8. \"What percentage of rows have Quantity greater than 1000?\""
    echo "9. \"What is the earliest ExDate in the dataset?\""
    echo "10. \"What is the TotalBaseIncome-to-Quantity ratio for SecurityID 828806109?\""
    echo "11. \"For which SecurityGroup is the variance of Quantity the highest, and what is that variance?\""
    echo "12. \"What is the difference between the median Quantity of DIV payments and the median Quantity of ACTUAL payments (DIV − ACTUAL)?\""
    echo "13. \"What is the ratio of the total TotalBaseIncome of ACTUAL rows to that of DIV rows?\""
    echo "14. \"What is the 90th percentile (P90) of Quantity across the dataset?\""
    echo "15. \"Which SecurityID shows the greatest standard deviation of TotalBaseIncome, and what is that standard deviation?\""
    echo "16. \"Which SecurityGroup derives the largest share of its TotalBaseIncome from DIV payments, and what is that share?\""
    echo "17. \"What is the average daily TotalBaseIncome over the full range of ExDates (earliest→latest, inclusive)?\""
    echo "18. \"How many rows record a negative TotalBaseIncome?\""
    echo "19. \"For SecurityID 101121101, what is the coefficient of variation of Quantity (std ÷ mean)?\""
    exit 0
fi

# Run the query directly (no build)
echo -e "🚀 Running query: $1\n"
./bin/Debug/net9.0/linux-x64/ssllm expanded_dataset_moved.xlsx "$1"