# Arixcel Explorer (Replica)

Windows desktop Excel add-in that replicates the core features of [Arixcel Explorer](https://www.arixcel.com/): formula logic exploration, dependents tracing, formula map, calculation flow, compare, and utility shortcuts.

Built as a **C# VSTO-style COM add-in** with a **VBA shortcut companion**, porting domain logic from the open-source [XLerate](https://github.com/omegarhovega/XLerate) project (MIT).

## Features

- **Explore Precedents** (`Ctrl+Q`) — logical formula tree with values and locations
- **Explore Dependents** (`Ctrl+Shift+Q`) — multi-cell direct dependents with Count column
- **Formula Map** — color-code inputs, formulas, links, inconsistencies, hardcodes
- **Calculation Flow** — highlight inputs, calculations, and outputs in a selection
- **Compare** — diff active worksheet against another sheet
- **Utility shortcuts** — unhide all sheets, select current region, toggle formula view
- **Options** — close behavior and large-scan thresholds

## Solution layout

```
Arixcel.sln
src/
  ArixcelExplorer/          VSTO add-in entry, ribbon, Excel interop services
  ArixcelExplorer.Core/     Pure C# domain (unit tested)
  ArixcelExplorer.UI/         WPF Explorer / Dependents / Options / Compare windows
  ArixcelShortcuts/         VBA .xlam source for keyboard shortcuts
tests/
  ArixcelExplorer.Core.Tests/
docs/
  project-context.md
  xl-erate-port-map.md
```

## Prerequisites

- Windows 10/11, Excel Desktop 64-bit (Microsoft 365 or Excel 2021+)
- Visual Studio 2022 with **Office/SharePoint development** workload
- [.NET Framework 4.8 Developer Pack](https://dotnet.microsoft.com/download/dotnet-framework/net48)
- [Visual Studio Tools for Office Runtime](https://learn.microsoft.com/en-us/visualstudio/vsto/visual-studio-tools-for-office-runtime-overview)

## Build (Windows)

```powershell
# Restore and build core tests (any OS with .NET SDK)
dotnet test tests/ArixcelExplorer.Core.Tests/ArixcelExplorer.Core.Tests.csproj

# Full solution (requires Windows + VS for WPF/VSTO/Excel interop)
msbuild Arixcel.sln /p:Configuration=Release /p:Platform="Any CPU"
```

> **Note:** The `ArixcelExplorer` project ships as a class library scaffold compatible with conversion to a full VSTO Excel Add-in in Visual Studio. Open the solution on Windows, use **Add > New Item > VSTO Add-in** migration or create a new Excel VSTO project and reference these projects.

## Install / sideload

1. Build `ArixcelExplorer` and register the COM add-in (VSTO publish or manual registry — see VS Publish wizard).
2. Import `src/ArixcelShortcuts/ArixcelShortcuts.bas` into an Excel `.xlam` add-in, or compile the companion add-in.
3. Enable both add-ins under **File → Options → Add-ins → COM Add-ins** and **Excel Add-ins**.
4. Confirm the **Arixcel** ribbon tab appears.

### VBA companion

The VBA module registers:

- `Ctrl+Q` → `OpenExplorer`
- `Ctrl+Shift+Q` → `OpenDependents`

It calls the COM add-in via `Application.COMAddIns("ArixcelExplorer.ComApi")`.

## Development on Core only (Linux/macOS)

The pure domain library builds and tests without Excel:

```bash
dotnet test tests/ArixcelExplorer.Core.Tests/ArixcelExplorer.Core.Tests.csproj
```

## License

Application code: project license TBD.

Ported XLerate logic: MIT (see [XLerate LICENSE](https://github.com/omegarhovega/XLerate/blob/master/LICENSE)).

Arixcel product name and UI are used as design reference only; this is an independent implementation.
