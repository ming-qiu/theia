#!/bin/bash
# Theia Installer for macOS
# Double-click to install

set -e

echo "======================================"
echo "Theia Pipeline Tools Installer"
echo "======================================"
echo ""

THEIA_DIR="/Library/Application Support/Theia"
RESOLVE_DIR="/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Edit"
INSTALLER_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Determine which Python to use
PYTHON_CMD="python3"

# Check for Homebrew Python (ARM64 native on Apple Silicon)
if [ -f "/opt/homebrew/bin/python3" ]; then
    echo "Found Homebrew Python (native ARM64)"
    PYTHON_CMD="/opt/homebrew/bin/python3"
elif [ -f "/usr/local/bin/python3" ]; then
    echo "Found Python in /usr/local/bin"
    PYTHON_CMD="/usr/local/bin/python3"
fi

# Check Python
echo "Checking Python installation..."
if ! command -v $PYTHON_CMD &> /dev/null; then
    echo "ERROR: Python 3 not found."
    echo ""
    echo "Please install Python 3.9+ from:"
    echo "https://www.python.org/downloads/"
    echo "OR install via Homebrew: brew install python@3.11"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

PYTHON_VERSION=$($PYTHON_CMD --version | cut -d' ' -f2)
PYTHON_MAJOR=$(echo $PYTHON_VERSION | cut -d'.' -f1)
PYTHON_MINOR=$(echo $PYTHON_VERSION | cut -d'.' -f2)

if [ "$PYTHON_MAJOR" -lt 3 ] || ([ "$PYTHON_MAJOR" -eq 3 ] && [ "$PYTHON_MINOR" -lt 9 ]); then
    echo "ERROR: Python $PYTHON_VERSION found, but 3.9+ required"
    exit 1
fi

echo "✓ Found Python $PYTHON_VERSION"
echo "  Using: $PYTHON_CMD"

# Check architecture - detect what architecture Python is actually running as
PYTHON_ARCH=$($PYTHON_CMD -c "import platform; print(platform.machine())")

echo "  Python running as: $PYTHON_ARCH"

# Use the runtime architecture to determine pip installation method
if [ "$PYTHON_ARCH" = "arm64" ]; then
    echo "  → Using native ARM64 mode"
    USE_ARCH_PREFIX=false
elif [ "$PYTHON_ARCH" = "x86_64" ]; then
    echo "  → Using x86_64 mode"
    USE_ARCH_PREFIX=true
else
    echo "  → Could not determine architecture, using default"
    USE_ARCH_PREFIX=false
fi

# Create directories
echo ""
echo "Creating Theia directory structure..."
echo "(This requires administrator password)"
sudo mkdir -p "$THEIA_DIR"
sudo chown -R $(whoami):staff "$THEIA_DIR"
sudo chmod -R 775 "$THEIA_DIR"
mkdir -p "$RESOLVE_DIR"

echo "✓ Directories created with proper permissions"

# Create virtual environment
echo ""
echo "Setting up Python environment..."
if [ -d "$THEIA_DIR/venv" ]; then
    echo "  Removing existing venv..."
    rm -rf "$THEIA_DIR/venv"
fi

python3 -m venv "$THEIA_DIR/venv"

if [ ! -f "$THEIA_DIR/venv/bin/activate" ]; then
    echo "ERROR: Failed to create virtual environment"
    exit 1
fi

echo "✓ Virtual environment created"

# Activate and install dependencies
echo ""
echo "Installing Python packages..."
source "$THEIA_DIR/venv/bin/activate"

pip install --upgrade pip setuptools wheel
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to upgrade pip"
    exit 1
fi

# Install packages with appropriate architecture handling
if [ "$USE_ARCH_PREFIX" = true ]; then
    echo "  Installing packages for x86_64 architecture (Rosetta mode)..."
    # Force x86_64 architecture for all pip installs
    arch -x86_64 pip install \
        PySide6 \
        "openpyxl>=3.0.0" \
        "Pillow>=8.0.0" \
        "timecode>=1.4.0" \
        "moviepy>=1.0.3"
else
    echo "  Installing packages for native architecture..."
    pip install \
        PySide6 \
        "openpyxl>=3.0.0" \
        "Pillow>=8.0.0" \
        "timecode>=1.4.0" \
        "moviepy>=1.0.3"
fi

if [ $? -ne 0 ]; then
    echo "ERROR: Failed to install dependencies"
    exit 1
fi

echo "✓ Dependencies installed"
deactivate

# Copy GUI scripts
echo ""
echo "Installing Theia tools..."

# Create resources directory
mkdir -p "$THEIA_DIR/resources"

# Create log directory
mkdir -p "$THEIA_DIR/log"

# Copy GUI scripts
for gui_script in "$INSTALLER_DIR"/scripts/*_gui.py; do
    if [ -f "$gui_script" ]; then
        cp "$gui_script" "$THEIA_DIR/"
        echo "  ✓ $(basename "$gui_script")"
    fi
done

# Copy resources (styles and icons)
if [ -d "$INSTALLER_DIR/resources" ]; then
    cp -R "$INSTALLER_DIR/resources/"* "$THEIA_DIR/resources/"
    echo "  ✓ Resources (styles, icons)"
fi

# Install bridge scripts
echo ""
echo "Installing Resolve bridge scripts..."
for bridge in "$INSTALLER_DIR"/bridges/*.py; do
    if [ -f "$bridge" ]; then
        cp "$bridge" "$RESOLVE_DIR/"
        chmod +x "$RESOLVE_DIR/$(basename "$bridge")"
        echo "  ✓ $(basename "$bridge")"
    fi
done

echo ""
echo "======================================"
echo "Installation Complete!"
echo "======================================"
echo ""
echo "Theia tools installed system-wide:"
echo "  Location: $THEIA_DIR"
echo "  Available to: All users"
echo ""
echo "Bridge scripts installed for current user:"
echo "  Location: $RESOLVE_DIR"
echo ""
echo "Note: Each user needs to run this installer once to"
echo "      set up their personal Resolve bridge scripts."
echo ""
echo "In DaVinci Resolve:"
echo "  Workspace → Scripts → Edit → [Tool Name]"
