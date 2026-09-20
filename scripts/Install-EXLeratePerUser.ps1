<#
.SYNOPSIS
  Per-user EXLerate install: copy payload to %LOCALAPPDATA%\EXLerate and register HKCU add-ins.
#>
param(
    [string]$PayloadDir,
    [string]$InstallDir,
    [switch]$SkipCopy,
    [switch]$SkipArp,
    [switch]$Silent
)

$ErrorActionPreference = 'Stop'
$exitCode = 0

function Write-Step([string]$Message, [string]$Color = 'Cyan') {
    if (-not $Silent) { Write-Host $Message -ForegroundColor $Color }
}

function Test-ExcelRunning {
    return [bool](Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue)
}

try {
    if ($SkipCopy) {
        if (-not $InstallDir) { $InstallDir = $PSScriptRoot }
    } else {
        if (-not $PayloadDir) { $PayloadDir = $PSScriptRoot }
        $PayloadDir = $PayloadDir.Trim().TrimEnd('\', '/')
        $PayloadDir = (Resolve-Path -LiteralPath $PayloadDir).Path
        if (-not $InstallDir) { $InstallDir = Join-Path $env:LOCALAPPDATA 'EXLerate' }
    }

    $registerScript = Join-Path $PSScriptRoot 'Register-ArixcelComAddIn.ps1'
    if (-not (Test-Path $registerScript)) {
        $registerScript = Join-Path $InstallDir 'Register-ArixcelComAddIn.ps1'
    }
    if (-not (Test-Path $registerScript) -and $PayloadDir) {
        $registerScript = Join-Path $PayloadDir 'Register-ArixcelComAddIn.ps1'
    }
    if (-not (Test-Path $registerScript)) {
        throw "Register-ArixcelComAddIn.ps1 not found next to this script."
    }

    if (Test-ExcelRunning) {
        Write-Step "Excel is running. Finish setup, then fully restart Excel to load EXLerate." 'Yellow'
    }

    if (-not $SkipCopy) {
        Write-Step "Installing to $InstallDir"
        New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null

        $copyExt = @('.dll', '.vsto', '.manifest', '.xlam', '.ps1', '.cmd', '.xml')
        Get-ChildItem -Path $PayloadDir -File | Where-Object {
            $copyExt -contains $_.Extension.ToLowerInvariant()
        } | ForEach-Object {
            Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $InstallDir $_.Name) -Force
        }

        $ribbonSrc = Join-Path $PayloadDir 'Ribbon'
        if (Test-Path $ribbonSrc) {
            $ribbonDest = Join-Path $InstallDir 'Ribbon'
            New-Item -ItemType Directory -Path $ribbonDest -Force | Out-Null
            Copy-Item -Path (Join-Path $ribbonSrc '*') -Destination $ribbonDest -Recurse -Force
        }
    }

    $vsto = Join-Path $InstallDir 'ArixcelExplorer.vsto'
    $dll = Join-Path $InstallDir 'ArixcelExplorer.dll'
    if (-not (Test-Path $vsto)) { throw "ArixcelExplorer.vsto not found in $InstallDir" }
    if (-not (Test-Path $dll)) { throw "ArixcelExplorer.dll not found in $InstallDir" }

    Write-Step "Registering EXLerate COM add-ins (HKCU)..."
    & $registerScript -VstoPath $vsto -DllPath $dll | Out-Null

    $xlam = Get-ChildItem -Path $InstallDir -Filter 'EXLerateShortcuts.xlam' -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($xlam) {
        $key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
        if (Test-Path $key) {
            $props = Get-ItemProperty -Path $key
            $keep = New-Object System.Collections.Generic.List[string]
            foreach ($name in @($props.PSObject.Properties.Name | Where-Object { $_ -match '^OPEN\d*$' } | Sort-Object { if ($_ -eq 'OPEN') { 0 } else { [int]($_ -replace '\D', '') } })) {
                $raw = [string]$props.$name
                foreach ($part in ($raw -split [char]0)) {
                    if ([string]::IsNullOrWhiteSpace($part)) { continue }
                    if ($part -like '*ArixcelExplorer.xlam*') { continue }
                    if ($part -like '*ArixcelShortcuts.xlam*') { continue }
                    if ($part -like '*EXLerateShortcuts.xlam*') { continue }
                    $keep.Add($part)
                }
                Remove-ItemProperty -Path $key -Name $name -ErrorAction SilentlyContinue
            }
            $keep.Add('"' + $xlam.FullName + '"')
            for ($i = 0; $i -lt $keep.Count; $i++) {
                $name = if ($i -eq 0) { 'OPEN' } else { "OPEN$i" }
                Set-ItemProperty -Path $key -Name $name -Value $keep[$i]
            }
            Write-Step "Registered Excel auto-open: $($xlam.FullName)" 'Green'
        }
    } else {
        Write-Step "EXLerateShortcuts.xlam not in payload - ribbon tab still works; Ctrl+Q needs the .xlam." 'Yellow'
    }

    if (-not $SkipArp) {
        $arp = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\EXLerateExplorer'
        New-Item -Path $arp -Force | Out-Null
        $uninstall = Join-Path $InstallDir 'Uninstall-EXLeratePerUser.ps1'
        Set-ItemProperty -Path $arp -Name 'DisplayName' -Value 'EXLerate Explorer'
        Set-ItemProperty -Path $arp -Name 'Publisher' -Value 'EXLerate'
        Set-ItemProperty -Path $arp -Name 'DisplayVersion' -Value '1.0.0'
        Set-ItemProperty -Path $arp -Name 'InstallLocation' -Value $InstallDir
        Set-ItemProperty -Path $arp -Name 'UninstallString' -Value ('powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' + $uninstall + '"')
        Set-ItemProperty -Path $arp -Name 'QuietUninstallString' -Value ('powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' + $uninstall + '" -Silent')
        New-ItemProperty -Path $arp -Name 'NoModify' -Value 1 -PropertyType DWord -Force | Out-Null
        New-ItemProperty -Path $arp -Name 'NoRepair' -Value 1 -PropertyType DWord -Force | Out-Null
    }

    Write-Step ""
    Write-Step "EXLerate installed for this Windows user." 'Green'
    Write-Step "Fully restart Excel. The EXLerate tab should appear."
    Write-Step "Install folder: $InstallDir"
}
catch {
    $exitCode = 1
    Write-Error $_
    throw
}
finally {
    exit $exitCode
}
