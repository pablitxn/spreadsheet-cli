#!/bin/bash

# Fast query script - skips build for quicker iterations
# Usage: ./query-fast.sh "your question here"

if [ -z "$1" ]; then
    echo "Usage: ./query-fast.sh \"your question here\""
    echo "Example: ./query-fast.sh \"What is the total revenue?\""
    exit 1
fi

# Run the query directly (no build)
echo -e "🚀 Running query: $1\n"
./bin/Debug/net9.0/linux-x64/ssllm "$1"