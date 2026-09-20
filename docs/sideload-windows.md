# Windows sideload guide

Full install for the EXLerate Explorer VSTO COM add-in and VBA shortcut companion.

## Prerequisites

| Component | Install |
|---|---|
| Excel Desktop 64-bit | Microsoft 365 or Excel 2021+ |
| Visual Studio 2022 | [Community](https://visualstudio.microsoft.com/) with **Office/SharePoint development** workload |
| .NET Framework 4.8 SDK | Included with VS workload above |
| VSTO Runtime | [Download](https://aka.ms/vstor) or `winget install Microsoft.VSTOR` |

## Option A — Developer PC (sideload from source)

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

## Option C — Office laptop (no admin)

Build the installer on a developer PC, then copy it to the laptop. No Visual Studio and no administrator rights on the laptop.

**Developer PC**

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\Publish-EXLerateInstaller.ps1
```

That writes:

- `dist\EXLerate-Installer\Install.cmd` — always produced
- `dist\EXLerateSetup.msi` — if the WiX CLI is available
- `dist\clickonce\setup.exe` — VSTO ClickOnce (ribbon only; use the MSI or Install.cmd for Ctrl+Q as well)

**Office laptop**

1. Close Excel completely (all workbooks).
2. Copy `EXLerateSetup.msi` **or** the `EXLerate-Installer` folder.
3. Double-click the MSI, or run `Install.cmd`.
4. Reopen Excel. The **EXLerate** tab should appear.

Prerequisites already on most office PCs that ran Arixcel: Excel Desktop 64-bit and the [VSTO Runtime](https://aka.ms/vstor). If the runtime is missing, install it once (that step may need admin / IT).

Uninstall with `Uninstall.cmd` or **Settings -> Apps -> EXLerate Explorer**.

## Enable add-ins in Excel

**File → Options → Add-ins**

| Manage dropdown | Add-in | Action |
|---|---|---|
| COM Add-ins → Go | EXLerate Explorer | Check |
| COM Add-ins → Go | EXLerate Explorer API | Check |
| Excel Add-ins → Go | EXLerateShortcuts | Browse → select `dist\EXLerateShortcuts.xlam` |

Confirm the **EXLerate** ribbon tab appears. Test **Ctrl+Q** and **Ctrl+Shift+Q**.

## Troubleshooting

- **Build error: Office targets missing** — Re-run VS Installer → modify → add *Office/SharePoint development*.
- **COM add-in disabled** — Excel may disable add-ins after errors. Re-enable via COM Add-ins dialog.
- **VBA "not loaded"** — Ensure both COM entries are checked and Excel was restarted after registration.
- **Trust VBA project access** — File → Options → Trust Center → Trust Center Settings → Macro Settings → check *Trust access to the VBA project object model* (required for automated `.xlam` creation).
