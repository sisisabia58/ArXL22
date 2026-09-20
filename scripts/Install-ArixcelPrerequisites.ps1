#Requires -RunAsAdministrator
<#
.SYNOPSIS
  Installs prerequisites for building and sideloading the EXLerate Explorer VSTO add-in.
#>
param(
    [switch]$SkipVsBuildTools
)

$ErrorActionPreference = 'Stop'

function Test-VstoRuntime {
    $keys = @(
        'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R'
    )
    foreach ($key in $keys) {
        if (Test-Path $key) { return $true }
    }
    return $false
}

function Get-VsMsBuildPath {
    $candidates = @(
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    )
    foreach ($path in $candidates) {
        if (Test-Path $path) { return $path }
    }
    return $null
}

Write-Host "=== EXLerate Explorer - prerequisite check ===" -ForegroundColor Cyan

if (-not (Test-VstoRuntime)) {
    Write-Host "Installing VSTO Runtime..." -ForegroundColor Yellow
    winget install --id Microsoft.VSTOR --accept-package-agreements --accept-source-agreements
} else {
    Write-Host "VSTO Runtime: OK" -ForegroundColor Green
}

$msbuild = Get-VsMsBuildPath
if ($msbuild) {
    Write-Host "Visual Studio MSBuild: $msbuild" -ForegroundColor Green
} elseif (-not $SkipVsBuildTools) {
    Write-Host "Installing Visual Studio 2022 Build Tools (Office/SharePoint workload)..." -ForegroundColor Yellow
    Write-Host "This may take 15-30 minutes." -ForegroundColor Yellow
    winget install --id Microsoft.VisualStudio.2022.BuildTools `
        --override "--wait --passive --add Microsoft.VisualStudio.Workload.OfficeBuildTools --includeRecommended" `
        --accept-package-agreements --accept-source-agreements
    $msbuild = Get-VsMsBuildPath
}

if (-not $msbuild) {
    Write-Error @"
Visual Studio 2022 with Office/SharePoint development tools was not found.

Install manually:
  1. Visual Studio 2022 Community/Professional, OR Build Tools
  2. Workload: Office/SharePoint development (includes VSTO build tools)
  3. Individual component: .NET Framework 4.8 SDK

Then re-run: .\scripts\Invoke-ArixcelSideload.ps1
"@
}

$officeTargets = Join-Path (Split-Path (Split-Path $msbuild)) "Microsoft\VisualStudio\v17.0\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets"
if (-not (Test-Path $officeTargets)) {
    Write-Error "Office VSTO build targets not found at:`n  $officeTargets`nRe-run VS Installer and add 'Office/SharePoint development build tools'."
}

Write-Host "Office VSTO targets: OK" -ForegroundColor Green
Write-Host "Prerequisites ready." -ForegroundColor Green
