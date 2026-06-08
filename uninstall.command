#!/bin/bash
# Theia Uninstaller for macOS
# Double-click to uninstall

echo "======================================"
echo "Theia Uninstaller"
echo "======================================"
echo ""
echo "This will remove:"
echo "  /Library/Application Support/Theia"
echo "  /Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Edit/Theia"
echo ""
read -p "Are you sure you want to uninstall Theia? [y/N] " confirm
echo ""

if [[ "$confirm" != "y" && "$confirm" != "Y" ]]; then
    echo "Uninstall cancelled."
    read -p "Press Enter to exit..."
    exit 0
fi

THEIA_DIR="/Library/Application Support/Theia"
RESOLVE_DIR="/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Edit/Theia"

# Remove main Theia directory
if [ -d "$THEIA_DIR" ]; then
    rm -rf "$THEIA_DIR"
    echo "✓ Removed $THEIA_DIR"
else
    echo "  (Skipped: $THEIA_DIR not found)"
fi

# Remove Resolve bridge scripts
if [ -d "$RESOLVE_DIR" ]; then
    rm -rf "$RESOLVE_DIR"
    echo "✓ Removed $RESOLVE_DIR"
else
    echo "  (Skipped: $RESOLVE_DIR not found)"
fi

echo ""
echo "======================================"
echo "Uninstall Complete"
echo "======================================"
echo ""
read -p "Press Enter to exit..."
