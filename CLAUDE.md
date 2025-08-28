# ViewPreviewTool - Critical Build Rules

## IMPORTANT: Revit Version Compatibility Rules

### Revit 2024
- **MUST use .NET Framework 4.8**
- **Compatible DLLs**: 
  - `ViewPreviewTool_v1.0_FINAL.dll`
  - `ViewPreviewTool_v1.1_FINAL_ZoomPan_900x700.dll`
  - Any DLL built with .NET Framework 4.8
- **DO NOT use**: 
  - Any DLL from `bin/Release/` (these are .NET 8.0)
  - Any DLL built with .NET Core/.NET 5+/.NET 8.0
- **Compiler**: Use `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`
- **Note**: UI features may be limited due to C# 5.0 syntax restrictions

### Revit 2025 & 2026
- **MUST use .NET 8.0**
- **Compatible DLLs**: 
  - `bin/Release/ViewPreviewTool.dll`
  - Any DLL built with modern .NET SDK
- **Features**: Full modern UI with all C# 10+ features
- **Compiler**: Use MSBuild with .NET 8.0 SDK

## Build Commands

### For Revit 2024:
```batch
# Use .NET Framework compiler
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /target:library /out:ViewPreviewTool_2024.dll ...
```

### For Revit 2025/2026:
```batch
# Use modern MSBuild
MSBuild.exe ViewPreviewTool.csproj /p:TargetFramework=net8.0-windows ...
```

## Installer Rules

### Revit 2024 Installer:
- ALWAYS use a .NET Framework 4.8 compatible DLL
- Installer name: `ViewPreviewTool_Setup_2024.exe`
- Expected size: ~32KB (smaller due to .NET Framework)

### Revit 2025/2026 Installer:
- ALWAYS use the modern .NET 8.0 DLL
- Installer name: `ViewPreviewTool_Setup_2025_2026.exe`
- Expected size: ~49KB (larger due to modern UI)

## Common Mistakes to Avoid

1. **NEVER use .NET 8.0 DLLs for Revit 2024** - It will fail silently
2. **NEVER mix framework versions** - Each Revit version has strict requirements
3. **Always check DLL size** - .NET Framework DLLs are typically smaller (16-32KB) while .NET 8.0 DLLs are larger (33-50KB)

## Testing Checklist

Before releasing any installer:
- [ ] Verify correct .NET version for target Revit version
- [ ] Check DLL file size matches expected range
- [ ] Test in actual Revit version
- [ ] Confirm add-in appears in ribbon
- [ ] Verify all features work (zoom, pan, etc.)

---
*This file is read by Claude Code to ensure consistent builds for different Revit versions*