# View Preview Tool for Revit 2025 (CLICK-based)

A Revit add-in that shows view previews when you CLICK the View Preview button after selecting a view in the Project Browser.

## Features

- **Click-based Preview**: Select a view in Project Browser, then CLICK the button to see preview (NOT hover)
- **View Information**: Displays view name in the preview window
- **Professional UI**: 700x800 pixel preview window with proper panels
- **Smart Positioning**: Preview appears to the LEFT of Project Browser
- **One-Click Operation**: Simple click-based workflow

## Installation

1. **Build the Project**:
   - Open the project in Visual Studio
   - Build in Debug or Release mode (x64 platform)
   - Or use the provided `build.bat` script

2. **Install the Add-in**:
   - Copy `ViewPreviewTool.dll` to a permanent location (e.g., `C:\ProgramData\Autodesk\Revit\Addins\2025\ViewPreviewTool\`)
   - Update the `<Assembly>` path in `ViewPreviewTool.addin` to point to your DLL location
   - Copy `ViewPreviewTool.addin` to `C:\ProgramData\Autodesk\Revit\Addins\2025\`

3. **Start Revit**:
   - Launch Revit 2025
   - Look for "View Preview" button in the Add-ins tab

## Usage

1. Select a view in the Project Browser by clicking on it
2. Click the "View Preview" button in the Add-ins tab
3. A preview window will appear on the LEFT side showing the selected view
4. Close the preview window when done
5. Repeat: Select view → Click button → See preview

## Requirements

- Revit 2025
- .NET Framework 4.8
- Visual Studio 2019 or later (for building)

## Technical Details

- The tool uses Revit's `ImageExportOptions` API to generate view thumbnails
- Previews are generated at 150 DPI for good quality vs performance balance
- Temporary image files are automatically cleaned up after use
- Only printable views are shown (excludes browser organization items)

## Troubleshooting

- If previews don't appear, ensure the feature is enabled via the ribbon button
- Some complex views may take longer to generate previews
- Views that cannot be printed will show a placeholder message
- Check Revit's journal file for any error messages if issues persist

## License

This tool is provided as-is for educational and productivity purposes.