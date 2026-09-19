# Windows sideload guide

Full install for the Arixcel Explorer VSTO COM add-in and VBA shortcut companion.

## Prerequisites

| Component | Install |
|---|---|
| Excel Desktop 64-bit | Microsoft 365 or Excel 2021+ |
| Visual Studio 2022 | [Community](https://visualstudio.microsoft.com/) with **Office/SharePoint development** workload |
| .NET Framework 4.8 SDK | Included with VS workload above |
| VSTO Runtime | [Download](https://aka.ms/vstor) or `winget install Microsoft.VSTOR` |

## Option A — Automated (recommended)

Open **PowerShell as Administrator** in the repo root:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\Install-ArixcelPrerequisites.ps1
.\scripts\Invoke-ArixcelSideload.ps1
```

Restart Excel. Enable add-ins under **File → Options → Add-ins** if any are unchecked.

## Option B — Visual Studio

1. Open `Arixcel.sln` in Visual Studio 2022.
2. Right-click **ArixcelExplorer** → **Properties** → confirm **Office Application = Excel**.
3. Build → **Build Solution** (Release recommended).
4. In PowerShell:

```powershell
.\scripts\Register-ArixcelComAddIn.ps1 `
  -VstoPath "src\ArixcelExplorer\bin\Release\ArixcelExplorer.vsto" `
  -DllPath "src\ArixcelExplorer\bin\Release\ArixcelExplorer.dll"

.\scripts\New-ArixcelShortcutsXlam.ps1
```

5. Restart Excel and enable add-ins (see below).

## Enable add-ins in Excel

**File → Options → Add-ins**

| Manage dropdown | Add-in | Action |
|---|---|---|
| COM Add-ins → Go | Arixcel Explorer | Check |
| COM Add-ins → Go | Arixcel Explorer API | Check |
| Excel Add-ins → Go | ArixcelShortcuts | Browse → select `dist\ArixcelShortcuts.xlam` |

Confirm the **Arixcel** ribbon tab appears. Test **Ctrl+Q** and **Ctrl+Shift+Q**.

## Troubleshooting

- **Build error: Office targets missing** — Re-run VS Installer → modify → add *Office/SharePoint development*.
- **COM add-in disabled** — Excel may disable add-ins after errors. Re-enable via COM Add-ins dialog.
- **VBA "not loaded"** — Ensure both COM entries are checked and Excel was restarted after registration.
- **Trust VBA project access** — File → Options → Trust Center → Trust Center Settings → Macro Settings → check *Trust access to the VBA project object model* (required for automated `.xlam` creation).
