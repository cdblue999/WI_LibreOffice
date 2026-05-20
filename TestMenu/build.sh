#!/bin/bash
# Build TestMenu.oxt extension
cd "$(dirname "$0")"
echo "Building TestMenu.oxt..."
zip -r ../TestMenu.oxt META-INF Addons.xcu description.xml Scripts/
echo "Done! Extension saved to ../TestMenu.oxt"
