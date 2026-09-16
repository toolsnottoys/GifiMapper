# GIFI Mapper — Setup & Build

## Requirements
- Visual Studio 2022 (any edition, including Community — free)
- .NET 8 SDK (installed automatically with VS2022)
- NuGet internet access to restore ClosedXML (one-time, automatic)

## How to open in Visual Studio
1. Unzip this folder anywhere on your PC.
2. Double-click **GifiMapper.csproj** — Visual Studio opens it.
3. Press **F5** (or Ctrl+F5) to build and run.
   - Visual Studio will restore the ClosedXML NuGet package automatically.

## How to build from command line
```
cd GifiMapper
dotnet build -c Release
dotnet run
```

## Auto-loading the GIFI mapping
**IG_gifi_codes.csv** is included in this folder.
Copy it next to the built .exe (`bin\Release\net8.0-windows\`) and the
program will load it automatically on startup.

## Features
| Feature | Details |
|---|---|
| Browse + Load | Separate buttons for GIFI mapping and Trial Balance CSVs |
| Preview tabs | Four tabs: GIFI Map · Trial Balance · Merged Results · Unmatched |
| Merge | Left-joins Trial Balance onto GIFI map by Accounting Code |
| Unmatched highlight | Rows with no GIFI code are shaded amber in the Merged tab |
| Live search | Filter rows in the active tab as you type |
| Save Merged CSV | Full merged result with GIFI Code + GIFI Description columns added |
| Save Unmatched CSV | Only rows that had no matching GIFI code |
| Export Excel (.xlsx) | Three-sheet workbook: Merged / Unmatched / GIFI Map, with auto-filter + frozen header |
| Keyboard shortcuts | Ctrl+G load GIFI · Ctrl+T load TB · Ctrl+S save · Ctrl+E export Excel |
| Ctrl+C copy | Copies selected rows from any grid to clipboard (tab-separated) |
| Double-click unmatched | Shows a pop-up explaining the missing code |
