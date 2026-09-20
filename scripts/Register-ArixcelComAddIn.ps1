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

$VstoPath = (Resolve-Path $VstoPath).Path
$DllPath = (Resolve-Path $DllPath).Path
$vstoUri = ([Uri]$VstoPath).AbsoluteUri
if ([string]::IsNullOrWhiteSpace($vstoUri) -or $vstoUri -eq '|vstolocal') {
    $vstoUri = ([Uri](New-Object IO.FileInfo $VstoPath).FullName).AbsoluteUri
}
$comApiClsid = '{B2C3D4E5-F6A7-8901-BCDE-F12345678901}'
$regAsm = Join-Path ${env:WINDIR} 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
if (-not (Test-Path $regAsm)) {
    $regAsm = Join-Path ${env:WINDIR} 'Microsoft.NET\Framework\v4.0.30319\RegAsm.exe'
}

Write-Host "Registering COM classes..." -ForegroundColor Cyan
$regFile = Join-Path (Split-Path $DllPath) 'EXLerateExplorer.reg'
$userRegFile = Join-Path (Split-Path $DllPath) 'EXLerateExplorer_User.reg'
$oldEAP = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
& $regAsm /codebase $DllPath /regfile:$regFile 2>$null | Out-Null
$ErrorActionPreference = $oldEAP
if (Test-Path $regFile) {
    $content = Get-Content $regFile -Raw
    $userContent = $content -replace 'HKEY_CLASSES_ROOT\\', 'HKEY_CURRENT_USER\Software\Classes\'
    Set-Content -Path $userRegFile -Value $userContent -Encoding ASCII
    $oldEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    & reg import $userRegFile *>$null
    $ErrorActionPreference = $oldEAP
}

$clsidKey = "HKCU:\Software\Classes\CLSID\$comApiClsid\Implemented Categories\{000C0396-0000-0000-C000-000000000046}"
New-Item -Path $clsidKey -Force | Out-Null

foreach ($legacyKey in @(
    'HKCU:\Software\Microsoft\Office\Excel\Addins\ArixcelExplorer',
    'HKCU:\Software\Microsoft\Office\Excel\Addins\ArixcelExplorer.ComApi'
)) {
    if (Test-Path $legacyKey) {
        Remove-Item -Path $legacyKey -Recurse -Force
    }
}

# Main VSTO add-in (ribbon + WPF UI)
$vstoKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\EXLerateExplorer'
New-Item -Path $vstoKey -Force | Out-Null
Set-ItemProperty -Path $vstoKey -Name 'Description' -Value 'EXLerate Explorer formula auditing add-in'
Set-ItemProperty -Path $vstoKey -Name 'FriendlyName' -Value 'EXLerate Explorer'
Set-ItemProperty -Path $vstoKey -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $vstoKey -Name 'Manifest' -Value "$vstoUri|vstolocal"

$doNotDisable = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList'
New-Item -Path $doNotDisable -Force | Out-Null
New-ItemProperty -Path $doNotDisable -Name 'EXLerateExplorer' -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $doNotDisable -Name 'EXLerateExplorer.ComApi' -Value 1 -PropertyType DWord -Force | Out-Null

foreach ($disabledKey in @(
    'HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems',
    'HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\NotificationDisabledItems'
)) {
    if (Test-Path $disabledKey) {
        Remove-Item -Path $disabledKey -Recurse -Force
    }
}

# Trust the sideload folder so VSTO is not blocked on a fresh workbook.
$inclusionRoot = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
New-Item -Path $inclusionRoot -Force | Out-Null
$vstoText = Get-Content -Path $VstoPath -Raw
$publicKey = $null
if ($vstoText -match '(?s)(<RSAKeyValue>.*?</RSAKeyValue>)') {
    $publicKey = $Matches[1]
}
Get-ChildItem $inclusionRoot -ErrorAction SilentlyContinue | ForEach-Object {
    $url = (Get-ItemProperty $_.PSPath).Url
    if ($url -eq $vstoUri) {
        Remove-Item $_.PSPath -Recurse -Force
    }
}
if ($publicKey) {
    $inclusionKey = Join-Path $inclusionRoot ([guid]::NewGuid().ToString('B'))
    New-Item -Path $inclusionKey -Force | Out-Null
    New-ItemProperty -Path $inclusionKey -Name 'Url' -Value $vstoUri -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $inclusionKey -Name 'PublicKey' -Value $publicKey -PropertyType String -Force | Out-Null
}

$trusted = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationEXLerate'
New-Item -Path $trusted -Force | Out-Null
Set-ItemProperty -Path $trusted -Name 'Path' -Value ((Split-Path $VstoPath) + '\')
Set-ItemProperty -Path $trusted -Name 'AllowSubfolders' -Value 1 -Type DWord
Set-ItemProperty -Path $trusted -Name 'Description' -Value 'EXLerate Explorer sideload'

# VBA shortcut bridge (Ctrl+Q / Ctrl+Shift+Q)
$apiKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\EXLerateExplorer.ComApi'
New-Item -Path $apiKey -Force | Out-Null
Set-ItemProperty -Path $apiKey -Name 'Description' -Value 'EXLerate Explorer VBA API bridge'
Set-ItemProperty -Path $apiKey -Name 'FriendlyName' -Value 'EXLerate Explorer API'
Set-ItemProperty -Path $apiKey -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $apiKey -Name 'CommandLineSafe' -Value 0 -Type DWord
if (Get-ItemProperty -Path $apiKey -Name 'Connect' -ErrorAction SilentlyContinue) {
    Remove-ItemProperty -Path $apiKey -Name 'Connect' -ErrorAction SilentlyContinue
}

Write-Host "Registered:" -ForegroundColor Green
Write-Host "  VSTO add-in:  $vstoKey"
Write-Host "  VBA ComApi:   $apiKey"
Write-Host ""
Write-Host "Restart Excel, then verify under File -> Options -> Add-ins:" -ForegroundColor Yellow
Write-Host "  COM Add-ins: EXLerate Explorer + EXLerate Explorer API (both checked)"
Write-Host "  Excel Add-ins: EXLerateShortcuts (after running New-ArixcelShortcutsXlam.ps1)"
