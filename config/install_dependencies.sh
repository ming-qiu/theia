#!/bin/bash

# Installation script for DaVinci Resolve VFX Pipeline Scripts
# This script installs required Python dependencies using your system Python

echo "================================================"
echo "DaVinci Resolve VFX Editor Scripts - Dependency Installer"
echo "================================================"
echo ""

# Check if python3 is available
if ! command -v python3 &> /dev/null; then
    echo "Error: python3 not found in PATH"
    echo ""
    echo "Please install Python 3.6 or later and try again."
    echo "Visit: https://www.python.org/downloads/"
    exit 1
fi

PYTHON_CMD="python3"

echo "Found Python at:"
which "$PYTHON_CMD"
echo ""

# Check Python version
echo "Python version:"
"$PYTHON_CMD" --version
echo ""

# Check minimum version (3.6)
PYTHON_VERSION=$("$PYTHON_CMD" -c 'import sys; print(".".join(map(str, sys.version_info[:2])))')
PYTHON_MAJOR=$(echo "$PYTHON_VERSION" | cut -d. -f1)
PYTHON_MINOR=$(echo "$PYTHON_VERSION" | cut -d. -f2)

if [ "$PYTHON_MAJOR" -lt 3 ] || ([ "$PYTHON_MAJOR" -eq 3 ] && [ "$PYTHON_MINOR" -lt 6 ]); then
    echo "Error: Python 3.6 or later is required"
    echo "Current version: $PYTHON_VERSION"
    exit 1
fi

# Install pip if needed
echo "Ensuring pip is installed..."
"$PYTHON_CMD" -m ensurepip --default-pip 2>/dev/null

# Upgrade pip
echo "Upgrading pip..."
"$PYTHON_CMD" -m pip install --upgrade pip

echo ""
echo "Installing dependencies from requirements.txt..."
echo ""

# Install requirements
"$PYTHON_CMD" -m pip install -r requirements.txt

if [ $? -eq 0 ]; then
    echo ""
    echo "================================================"
    echo "✓ Installation complete!"
    echo "================================================"
    echo ""
    echo "You can now run the scripts. For example:"
    echo "  python3 clip-inventory.py"
    echo "  python3 shot-list.py --output my_shot_list.xlsx"
    echo ""
else
    echo ""
    echo "================================================"
    echo "✗ Installation failed"
    echo "================================================"
    echo ""
    echo "Please try installing manually:"
    echo "  python3 -m pip install -r requirements.txt"
    echo ""
    exit 1
fi