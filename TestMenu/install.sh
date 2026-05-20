#!/bin/bash
# Install TestMenu.oxt extension
cd "$(dirname "$0")"
EXTENSION="$PWD/TestMenu.oxt"

if command -v unopkg &> /dev/null; then
    echo "Installing TestMenu extension..."
    unopkg add "$EXTENSION"
    echo "Done! Restart LibreOffice to see the Test menu."
else
    echo "unopkg not found. Install manually:"
    echo "1. Open LibreOffice"
    echo "2. Tools > Extension Manager"
    echo "3. Click 'Add' and select: $EXTENSION"
    echo "4. Restart LibreOffice"
fi
