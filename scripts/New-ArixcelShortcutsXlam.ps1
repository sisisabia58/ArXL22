<#
.SYNOPSIS
  Builds the ArixcelShortcuts.xlam companion add-in from the VBA source module.
#>
param(
    [string]$OutputPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'dist\ArixcelShortcuts.xlam')
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$basPath = Join-Path $repoRoot 'src\ArixcelShortcuts\ArixcelShortcuts.bas'
$outDir = Split-Path $OutputPath -Parent

if (-not (Test-Path $basPath)) { throw "VBA source not found: $basPath" }
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

function Register-XlamAddIn {
    param([string]$XlamPath)
    $key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
    $openValue = 'OPEN'
    if (Test-Path $key) {
        $existing = Get-ItemProperty -Path $key -Name $openValue -ErrorAction SilentlyContinue
        $entry = """$XlamPath"""
        if ($existing.$openValue) {
            if ($existing.$openValue -notlike "*$XlamPath*") {
                Set-ItemProperty -Path $key -Name $openValue -Value ($existing.$openValue + [char]0 + $entry)
            }
        } else {
            Set-ItemProperty -Path $key -Name $openValue -Value $entry
        }
    }
    Write-Host "Registered Excel add-in auto-open: $XlamPath" -ForegroundColor Green
}

if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

Write-Host "Creating ArixcelShortcuts.xlam..." -ForegroundColor Cyan

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    $workbook.IsAddin = $true

    # xlModule = 1
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = 'ArixcelShortcuts'

    $lines = Get-Content $basPath -Raw
    # Strip the Attribute VB_Name line - VBComponents.Name sets it
    $lines = $lines -replace '(?m)^Attribute VB_Name = ".*"\r?\n', ''
    $module.CodeModule.AddFromString($lines)

    # xlOpenXMLAddIn = 55
    $workbook.SaveAs($OutputPath, 55)
    Write-Host "Saved: $OutputPath" -ForegroundColor Green
}
finally {
    if ($workbook) { $workbook.Close($false) | Out-Null }
    if ($excel) {
        $excel.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

Register-XlamAddIn -XlamPath $OutputPath

Write-Host ""
Write-Host "Enable in Excel: File -> Options -> Add-ins -> Excel Add-ins -> Browse -> select .xlam" -ForegroundColor Yellow

return $OutputPath
