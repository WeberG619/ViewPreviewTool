# ViewPreviewTool - WORKING VERSIONS

This directory contains the **PROVEN WORKING** versions of ViewPreviewTool that have been tested and confirmed to work without issues.

## Working Versions

### Revit 2025/2026
- **DLL**: `DLLs/ViewPreviewTool_2025_2026_WORKING.dll`
- **Source**: From `ViewPreviewTool_v1.0_Distribution/Addin/ViewPreviewTool.dll` (Aug 26, 2025)
- **Size**: 30,720 bytes
- **Framework**: .NET 8.0
- **Features**: 
  - Click-based preview (no hover issues)
  - Preview window stays open (no flashing)
  - Clean, modern UI
  - Zoom/pan functionality

### Revit 2024
- **DLL**: `DLLs/ViewPreviewTool_2024_WORKING.dll`
- **Source**: From `ViewPreviewTool_v1.0_FINAL.dll` (Aug 26, 2025)
- **Size**: 16,896 bytes
- **Framework**: .NET Framework 4.8
- **Features**: Same as above but compatible with older framework

## Installers

### Ready-to-Use Installers
1. `Installers/ViewPreviewTool_v1.0_Setup_2025_2026.exe` (94 KB)
   - Contains the working 2025/2026 DLL
   - BIM Ops Studio icon included
   - Professional installer for website distribution

2. `Installers/ViewPreviewTool_v1.0_Setup_2024.exe` (81 KB)
   - Contains the working 2024 DLL
   - BIM Ops Studio icon included
   - Professional installer for website distribution

## IMPORTANT NOTES

**DO NOT** use any of these DLLs:
- ❌ `bin/Release/ViewPreviewTool.dll` - May have hover/flashing issues
- ❌ `ViewPreviewTool_2024_FINAL.dll` - Has issues
- ❌ Any DLL built after these working versions without thorough testing

**ALWAYS** use the DLLs in this WORKING_VERSIONS directory for:
- Creating new installers
- Deploying to users
- Website downloads

## How to Build New Installers

If you need to rebuild the installers, use these exact DLLs:

```batch
# For Revit 2025/2026
csc.exe /target:winexe /out:Setup_2025_2026.exe /win32icon:BIMOpsStudio.ico /resource:WORKING_VERSIONS/DLLs/ViewPreviewTool_2025_2026_WORKING.dll,ViewPreviewTool.dll ...

# For Revit 2024
csc.exe /target:winexe /out:Setup_2024.exe /win32icon:BIMOpsStudio.ico /resource:WORKING_VERSIONS/DLLs/ViewPreviewTool_2024_WORKING.dll,ViewPreviewTool.dll ...
```

## Version History

- **Aug 28, 2025 4:35 PM**: Identified and saved these working versions
- **Aug 26, 2025**: Original working DLLs created

---

**Remember**: These are the ONLY versions confirmed to work properly. Always use these!