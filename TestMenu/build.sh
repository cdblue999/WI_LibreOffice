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
# Build TestMenu.oxt extension
cd "$(dirname "$0")"
echo "Building WI_Menu.oxt..."
zip -r ../WI_Menu.oxt META-INF Addons.xcu description.xml Scripts/
echo "Done! Extension saved to ../WI_Menu.oxt"
