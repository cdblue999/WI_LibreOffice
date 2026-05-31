# Copyright (C) 2026 ZMS
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#!/bin/bash
# Install TestMenu.oxt extension
cd "$(dirname "$0")"
EXTENSION="$PWD/../WI_Menu.oxt"

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
