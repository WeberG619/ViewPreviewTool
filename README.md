# View Preview Tool for Revit 2025

A Revit add-in that shows view previews when hovering over view names in the Project Browser.

## Features

- **Hover Preview**: Shows a thumbnail preview of views when hovering over them in the Project Browser
- **View Information**: Displays view name and type in the preview
- **Toggle Control**: Ribbon button to enable/disable the preview feature
- **Smart Positioning**: Preview appears near cursor with automatic positioning
- **Performance**: 500ms delay before showing preview to avoid accidental triggers

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
   - Look for "Toggle Preview" button in the "View Tools" panel

## Usage

1. Click the "Toggle Preview" button in the ribbon to enable the feature
2. Hover over any view name in the Project Browser
3. After a short delay (500ms), a preview popup will appear
4. Move your cursor away or click to dismiss the preview
5. Click the toggle button again to disable the feature

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