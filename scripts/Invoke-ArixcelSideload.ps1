<#
.SYNOPSIS
  Full Windows sideload: build VSTO add-in, register COM, create .xlam companion.
.DESCRIPTION
  Run from an elevated PowerShell on Windows with Excel Desktop installed.

  Example:
    Set-ExecutionPolicy -Scope Process Bypass
    .\scripts\Invoke-ArixcelSideload.ps1

  To install prerequisites first (admin):
    .\scripts\Install-ArixcelPrerequisites.ps1
#>
param(
    [switch]$SkipPrerequisites,
    [switch]$SkipBuild,
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent

Write-Host @"

  Arixcel Explorer - Windows sideload
  ===================================
  1. Build VSTO COM add-in (ArixcelExplorer)
  2. Register for Excel COM Add-ins
  3. Create ArixcelShortcuts.xlam from VBA source
  4. Enable both in Excel (File -> Options -> Add-ins)

"@ -ForegroundColor Cyan

if (-not $SkipPrerequisites) {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($isAdmin) {
        & (Join-Path $PSScriptRoot 'Install-ArixcelPrerequisites.ps1')
    } else {
        Write-Host "Tip: Run as Administrator once to auto-install VSTO Runtime / Build Tools." -ForegroundColor Yellow
    }
}

$buildResult = $null
if (-not $SkipBuild) {
    $buildResult = & (Join-Path $PSScriptRoot 'Build-ArixcelVsto.ps1') -Configuration $Configuration
} else {
    $outputDir = Join-Path $repoRoot "src\ArixcelExplorer\bin\$Configuration"
    $buildResult = @{
        OutputDir = $outputDir
        DllPath   = Join-Path $outputDir 'ArixcelExplorer.dll'
        VstoPath  = Join-Path $outputDir 'ArixcelExplorer.vsto'
    }
}

& (Join-Path $PSScriptRoot 'Register-ArixcelComAddIn.ps1') `
    -VstoPath $buildResult.VstoPath `
    -DllPath $buildResult.DllPath

$xlamPath = & (Join-Path $PSScriptRoot 'New-ArixcelShortcutsXlam.ps1')

Write-Host @"

  Sideload complete
  =================
  Restart Excel (close all instances first).

  Verify:
    File -> Options -> Add-ins
      Manage COM Add-ins -> Go:
        [x] Arixcel Explorer
        [x] Arixcel Explorer API
      Manage Excel Add-ins -> Go:
        [x] ArixcelShortcuts

  Confirm the "Arixcel" ribbon tab appears.
  Test Ctrl+Q (Explore Precedents) and Ctrl+Shift+Q (Explore Dependents).

  Artifacts:
    VSTO: $($buildResult.VstoPath)
    XLAM: $xlamPath

"@ -ForegroundColor Green
