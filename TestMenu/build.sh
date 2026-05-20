#!/bin/bash
# Build TestMenu.oxt extension
cd "$(dirname "$0")"
echo "Building WI_Menu.oxt..."
zip -r ../WI_Menu.oxt META-INF Addons.xcu description.xml Scripts/
echo "Done! Extension saved to ../WI_Menu.oxt"
