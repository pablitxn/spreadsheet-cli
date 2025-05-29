#!/bin/bash

echo "Testing audit log functionality..."
echo ""

# Run a simple query
./bin/Debug/net9.0/linux-x64/ssllm query -f expanded_dataset_moved.xlsx -q "What is the standard deviation of the Quantity column?" > /dev/null 2>&1

# Find the most recent audit log file
LATEST_LOG=$(ls -t audit_log_*.txt 2>/dev/null | head -1)

if [ -z "$LATEST_LOG" ]; then
    echo "❌ No audit log file found!"
    exit 1
fi

echo "✅ Audit log created: $LATEST_LOG"
echo ""
echo "📊 Log file size: $(wc -c < "$LATEST_LOG") bytes"
echo "📊 Total lines: $(wc -l < "$LATEST_LOG")"
echo ""
echo "🔍 Sample content (first 50 lines):"
echo "=================================="
head -50 "$LATEST_LOG"
echo ""
echo "=================================="
echo ""
echo "💡 Full log saved in: $PWD/$LATEST_LOG"
echo "💡 You can copy this file for audit purposes"