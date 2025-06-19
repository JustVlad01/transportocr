#!/bin/bash

# Transport Sorter - PDF Scanner
# Activation script to run the application

echo "🚀 Starting Transport Sorter - PDF Scanner..."

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "❌ Virtual environment not found. Please run setup.py first."
    exit 1
fi

# Activate virtual environment and run the application
source venv/bin/activate
echo "✅ Virtual environment activated"
echo "🎯 Launching application..."
python3 main.py