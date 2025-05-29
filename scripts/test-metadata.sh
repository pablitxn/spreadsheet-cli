#!/bin/bash

# Test the enhanced metadata extraction with standard deviation queries

echo "Testing enhanced metadata extraction..."
echo ""

# Test 1: Standard deviation query
echo "Test 1: Standard deviation calculation"
./bin/Debug/net9.0/linux-x64/ssllm query -f expanded_dataset_moved.xlsx -q "What is the standard deviation of the Quantity column?"
echo ""
echo "---"
echo ""

# Test 2: Variance query
echo "Test 2: Variance calculation"
./bin/Debug/net9.0/linux-x64/ssllm query -f expanded_dataset_moved.xlsx -q "Calculate the variance of the Quantity values"
echo ""
echo "---"
echo ""

# Test 3: Multiple statistics query
echo "Test 3: Multiple statistics"
./bin/Debug/net9.0/linux-x64/ssllm query -f expanded_dataset_moved.xlsx -q "Give me the mean, median, and standard deviation of the Price column"
echo ""