# Test Menu Extension for LibreOffice

Adds a **Test** menu group to the LibreOffice main menu bar.

## Menu Items
- **Hello World** - Shows a message box
- **Insert Text** - Inserts text at cursor position
- **Set Cell Color** - Colors cell A1 yellow (Calc only)

## Install

### Automatic
```bash
chmod +x install.sh && ./install.sh
```

### Manual
1. Open LibreOffice
2. **Tools** > **Extension Manager**
3. Click **Add** and select `TestMenu.oxt`
4. Restart LibreOffice

## Build
```bash
./build.sh
```

## Structure
```
TestMenu/
├── META-INF/manifest.xml    # Extension manifest
├── Addons.xcu               # Menu configuration
├── description.xml          # Extension metadata
├── Scripts/basic/
│   ├── Module1.xba          # Basic macros
│   └── script.xlb           # Library definition
├── build.sh                 # Build script
└── install.sh               # Install script
```
