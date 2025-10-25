#!/bin/bash

# Installation script for DaVinci Resolve VFX Pipeline Scripts
# This script installs required Python dependencies using Resolve's Python

echo "================================================"
echo "DaVinci Resolve VFX Scripts - Dependency Installer"
echo "================================================"
echo ""

# Detect OS and set Resolve Python path
if [[ "$OSTYPE" == "darwin"* ]]; then
    # macOS
    RESOLVE_PYTHON="/Applications/DaVinci Resolve/DaVinci Resolve.app/Contents/Libraries/Frameworks/Python.framework/Versions/3.6/bin/python3"
    
    # Check for alternate locations
    if [ ! -f "$RESOLVE_PYTHON" ]; then
        RESOLVE_PYTHON="/Applications/DaVinci Resolve/DaVinci Resolve.app/Contents/Libraries/Frameworks/Python.framework/Versions/Current/bin/python3"
    fi
    
elif [[ "$OSTYPE" == "linux-gnu"* ]]; then
    # Linux
    RESOLVE_PYTHON="/opt/resolve/bin/python"
    
    # Check alternate location
    if [ ! -f "$RESOLVE_PYTHON" ]; then
        RESOLVE_PYTHON="/opt/resolve/bin/python3"
    fi
else
    echo "Error: Unsupported OS. Please run install_dependencies.bat on Windows."
    exit 1
fi

# Verify Resolve Python exists
if [ ! -f "$RESOLVE_PYTHON" ]; then
    echo "Error: Could not find DaVinci Resolve Python at:"
    echo "  $RESOLVE_PYTHON"
    echo ""
    echo "Please locate your Resolve Python installation and run:"
    echo "  /path/to/resolve/python3 -m pip install -r requirements.txt"
    exit 1
fi

echo "Found DaVinci Resolve Python at:"
echo "  $RESOLVE_PYTHON"
echo ""

# Check Python version
echo "Python version:"
"$RESOLVE_PYTHON" --version
echo ""

# Install pip if needed
echo "Ensuring pip is installed..."
"$RESOLVE_PYTHON" -m ensurepip --default-pip 2>/dev/null

# Upgrade pip
echo "Upgrading pip..."
"$RESOLVE_PYTHON" -m pip install --upgrade pip

echo ""
echo "Installing dependencies from requirements.txt..."
echo ""

# Install requirements
"$RESOLVE_PYTHON" -m pip install -r requirements.txt

if [ $? -eq 0 ]; then
    echo ""
    echo "================================================"
    echo "✓ Installation complete!"
    echo "================================================"
    echo ""
    echo "You can now run the scripts. For example:"
    echo "  python clip-inventory.py"
    echo "  python shot-list.py --output my_shot_list.xlsx"
    echo ""
else
    echo ""
    echo "================================================"
    echo "✗ Installation failed"
    echo "================================================"
    echo ""
    echo "Please try installing manually:"
    echo "  \"$RESOLVE_PYTHON\" -m pip install -r requirements.txt"
    echo ""
    exit 1
fi
