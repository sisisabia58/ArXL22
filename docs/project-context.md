# EXLerate Explorer — Project Context

## Goal

Ship **EXLerate Explorer** for **Excel Desktop on Windows**: formula logic exploration, multi-cell dependents tracing, formula map, calculation flow, compare, and utility shortcuts. Design reference: [Arixcel Explorer](https://www.arixcel.com/).

## Architecture

- **ArixcelExplorer** — VSTO-style COM add-in (C# / .NET Framework 4.8) with custom ribbon
- **ArixcelExplorer.Core** — pure domain logic (ported from [XLerate](https://github.com/omegarhovega/XLerate) TypeScript core)
- **ArixcelExplorer.UI** — WPF windows for EXLerate Explorer layout
- **ArixcelShortcuts** — VBA `.xlam` companion (`EXLerateShortcuts`) for `Ctrl+Q` / `Ctrl+Shift+Q`

## Feature parity matrix

| Feature | Status | Notes |
|---|---|---|
| Explore Precedents (`Ctrl+Q`) | Implemented | Formula AST tree + WPF Explorer window |
| Explore Dependents (`Ctrl+Shift+Q`) | Implemented | Multi-cell, Count column, keyboard nav |
| Formula Map | Implemented | autoColor + horizontal consistency |
| Calculation Flow | Implemented | Input / calculation / output overlay |
| Compare | Implemented | Active sheet vs first other sheet (MVP) |
| Utility shortcuts | Implemented | Unhide sheets, select region, toggle formulas |
| Options | Implemented | Close behavior, scan thresholds |
| Workbook compare / alignment | Partial | Row alignment helper in Core; full UI later |

## Development requirements

- Windows 10/11
- Microsoft 365 or Excel 2021+ Desktop (64-bit)
- Visual Studio 2022 with **Office/SharePoint development** workload
- [VSTO Runtime](https://learn.microsoft.com/en-us/visualstudio/vsto/visual-studio-tools-for-office-runtime-overview)

## Keyboard shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl+Q` | Explore Precedents |
| `Ctrl+Shift+Q` | Explore Dependents |
| `Ctrl+Backspace` | Return to origin (explorer stack) |

## XLerate attribution

Logic ported from XLerate (MIT License). See [xl-erate-port-map.md](./xl-erate-port-map.md).
