<#
.SYNOPSIS
  Unregisters EXLerate HKCU add-ins and removes %LOCALAPPDATA%\EXLerate.
#>
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'EXLerate'),
    [switch]$KeepFiles,
    [switch]$Silent
)

$ErrorActionPreference = 'Stop'

function Write-Step([string]$Message, [string]$Color = 'Cyan') {
    if (-not $Silent) { Write-Host $Message -ForegroundColor $Color }
}

if (Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue) {
    throw "Close Excel completely before uninstalling EXLerate."
}

Write-Step "Unregistering EXLerate..."

foreach ($key in @(
    'HKCU:\Software\Microsoft\Office\Excel\Addins\EXLerateExplorer',
    'HKCU:\Software\Microsoft\Office\Excel\Addins\EXLerateExplorer.ComApi',
    'HKCU:\Software\Microsoft\Office\Excel\Addins\ArixcelExplorer',
    'HKCU:\Software\Microsoft\Office\Excel\Addins\ArixcelExplorer.ComApi',
    'HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationEXLerate'
)) {
    if (Test-Path $key) { Remove-Item -Path $key -Recurse -Force }
}

$doNotDisable = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList'
if (Test-Path $doNotDisable) {
    foreach ($name in @('EXLerateExplorer', 'EXLerateExplorer.ComApi')) {
        Remove-ItemProperty -Path $doNotDisable -Name $name -ErrorAction SilentlyContinue
    }
}

$inclusionRoot = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
if (Test-Path $inclusionRoot) {
    $needle = 'ArixcelExplorer.vsto'
    Get-ChildItem $inclusionRoot -ErrorAction SilentlyContinue | ForEach-Object {
        $url = [string](Get-ItemProperty $_.PSPath).Url
        if ($url -like "*$needle*") {
            Remove-Item $_.PSPath -Recurse -Force
        }
    }
}

$optKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
if (Test-Path $optKey) {
    $props = Get-ItemProperty -Path $optKey
    $keep = New-Object System.Collections.Generic.List[string]
    foreach ($name in @($props.PSObject.Properties.Name | Where-Object { $_ -match '^OPEN\d*$' } | Sort-Object { if ($_ -eq 'OPEN') { 0 } else { [int]($_ -replace '\D', '') } })) {
        $raw = [string]$props.$name
        foreach ($part in ($raw -split [char]0)) {
            if ([string]::IsNullOrWhiteSpace($part)) { continue }
            if ($part -like '*EXLerateShortcuts.xlam*') { continue }
            if ($part -like '*ArixcelShortcuts.xlam*') { continue }
            if ($part -like '*ArixcelExplorer.xlam*') { continue }
            $keep.Add($part)
        }
        Remove-ItemProperty -Path $optKey -Name $name -ErrorAction SilentlyContinue
    }
    for ($i = 0; $i -lt $keep.Count; $i++) {
        $name = if ($i -eq 0) { 'OPEN' } else { "OPEN$i" }
        Set-ItemProperty -Path $optKey -Name $name -Value $keep[$i]
    }
}

$comApiClsid = '{B2C3D4E5-F6A7-8901-BCDE-F12345678901}'
foreach ($prog in @(
    'HKCU:\Software\Classes\EXLerateExplorer.ComApi',
    "HKCU:\Software\Classes\CLSID\$comApiClsid"
)) {
    if (Test-Path $prog) { Remove-Item -Path $prog -Recurse -Force }
}

$arp = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\EXLerateExplorer'
if (Test-Path $arp) { Remove-Item -Path $arp -Recurse -Force }

if (-not $KeepFiles -and (Test-Path $InstallDir)) {
    Remove-Item -Path $InstallDir -Recurse -Force
    Write-Step "Removed $InstallDir" 'Green'
}

Write-Step "EXLerate uninstalled for this Windows user." 'Green'
