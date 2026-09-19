<#
.SYNOPSIS
  Registers the VSTO COM add-in and the VBA-facing ComApi companion for Excel sideload.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$VstoPath,

    [Parameter(Mandatory = $true)]
    [string]$DllPath
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $VstoPath)) { throw "VSTO manifest not found: $VstoPath" }
if (-not (Test-Path $DllPath)) { throw "Assembly not found: $DllPath" }

$vstoUri = ([Uri]$VstoPath).AbsoluteUri
$comApiClsid = '{B2C3D4E5-F6A7-8901-BCDE-F12345678901}'
$regAsm = Join-Path ${env:WINDIR} 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
if (-not (Test-Path $regAsm)) {
    $regAsm = Join-Path ${env:WINDIR} 'Microsoft.NET\Framework\v4.0.30319\RegAsm.exe'
}

Write-Host "Registering COM classes..." -ForegroundColor Cyan
& $regAsm /codebase $DllPath /tlb:"$($DllPath -replace '\.dll$','.tlb')" 2>&1 | Out-Null

# Main VSTO add-in (ribbon + WPF UI)
$vstoKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\ArixcelExplorer'
New-Item -Path $vstoKey -Force | Out-Null
Set-ItemProperty -Path $vstoKey -Name 'Description' -Value 'Arixcel Explorer formula auditing add-in'
Set-ItemProperty -Path $vstoKey -Name 'FriendlyName' -Value 'Arixcel Explorer'
Set-ItemProperty -Path $vstoKey -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $vstoKey -Name 'Manifest' -Value "$vstoUri|vstolocal"

# VBA shortcut bridge (Ctrl+Q / Ctrl+Shift+Q)
$apiKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\ArixcelExplorer.ComApi'
New-Item -Path $apiKey -Force | Out-Null
Set-ItemProperty -Path $apiKey -Name 'Description' -Value 'Arixcel Explorer VBA API bridge'
Set-ItemProperty -Path $apiKey -Name 'FriendlyName' -Value 'Arixcel Explorer API'
Set-ItemProperty -Path $apiKey -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $apiKey -Name 'CommandLineSafe' -Value 0 -Type DWord
Set-ItemProperty -Path $apiKey -Name 'Connect' -Value $comApiClsid

Write-Host "Registered:" -ForegroundColor Green
Write-Host "  VSTO add-in:  $vstoKey"
Write-Host "  VBA ComApi:   $apiKey"
Write-Host ""
Write-Host "Restart Excel, then verify under File -> Options -> Add-ins:" -ForegroundColor Yellow
Write-Host "  COM Add-ins: Arixcel Explorer + Arixcel Explorer API (both checked)"
Write-Host "  Excel Add-ins: ArixcelShortcuts (after running New-ArixcelShortcutsXlam.ps1)"
