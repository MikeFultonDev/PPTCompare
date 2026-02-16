#!/bin/bash

echo "PowerPoint Comparison Tool - Setup Script"
echo "=========================================="
echo ""

# Check if running on macOS
if [[ "$OSTYPE" == "darwin"* ]]; then
    echo "Detected macOS"
    
    # Check for Homebrew
    if ! command -v brew &> /dev/null; then
        echo "Error: Homebrew not found. Please install Homebrew first:"
        echo "  /bin/bash -c \"\$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\""
        exit 1
    fi
    
    # Check for LibreOffice
    if ! command -v libreoffice &> /dev/null && ! command -v soffice &> /dev/null; then
        echo "Installing LibreOffice..."
        brew install --cask libreoffice
    else
        echo "✓ LibreOffice already installed"
    fi
    
    # Check for Poppler
    if ! command -v pdfinfo &> /dev/null; then
        echo "Installing Poppler..."
        brew install poppler
    else
        echo "✓ Poppler already installed"
    fi
fi

# Create virtual environment
if [ ! -d "venv" ]; then
    echo ""
    echo "Creating virtual environment..."
    python3 -m venv venv
else
    echo "✓ Virtual environment already exists"
fi

# Activate virtual environment
echo ""
echo "Activating virtual environment..."
source venv/bin/activate

# Install Python dependencies
echo ""
echo "Installing Python dependencies..."
pip install --upgrade pip
pip install -r requirements.txt

echo ""
echo "=========================================="
echo "Setup complete!"
echo ""
echo "To use the tool:"
echo "  1. Activate the virtual environment: source venv/bin/activate"
echo "  2. Run: python ppt_compare.py <file1.pptx> <file2.pptx>"
echo ""

# Made with Bob
