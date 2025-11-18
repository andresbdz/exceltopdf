#!/bin/bash

# Excel to PDF Web Converter - Quick Start Script

echo "=========================================="
echo "Excel to PDF Converter - Web Application"
echo "=========================================="
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed. Please install Python 3 first."
    exit 1
fi

echo "✓ Python 3 found"

# Check if pip is installed
if ! command -v pip3 &> /dev/null; then
    echo "❌ pip is not installed. Please install pip first."
    exit 1
fi

echo "✓ pip found"

# Install dependencies
echo ""
echo "📦 Installing dependencies..."
pip3 install -q flask pandas reportlab openpyxl

if [ $? -eq 0 ]; then
    echo "✓ Dependencies installed successfully"
else
    echo "❌ Failed to install dependencies"
    exit 1
fi

# Start the server
echo ""
echo "🚀 Starting web server..."
echo ""
echo "=========================================="
echo "✓ Server is running!"
echo "=========================================="
echo ""
echo "📱 Open your browser and go to:"
echo ""
echo "   👉  http://localhost:5000"
echo ""
echo "=========================================="
echo ""
echo "Press Ctrl+C to stop the server"
echo ""

# Run the Flask app
python3 app.py
