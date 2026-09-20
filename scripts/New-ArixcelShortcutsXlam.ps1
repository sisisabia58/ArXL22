<#
.SYNOPSIS
  Builds the EXLerateShortcuts.xlam companion add-in from the VBA source module.
#>
param(
    [string]$OutputPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'dist\EXLerateShortcuts.xlam')
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$basPath = Join-Path $repoRoot 'src\ArixcelShortcuts\EXLerateShortcuts.bas'
$outDir = Split-Path $OutputPath -Parent

if (-not (Test-Path $basPath)) { throw "VBA source not found: $basPath" }
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

function Register-XlamAddIn {
    param([string]$XlamPath)
    $key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
    if (-not (Test-Path $key)) {
        Write-Host "Excel Options key missing; skip auto-open registration." -ForegroundColor Yellow
        return
    }

    $props = Get-ItemProperty -Path $key
    $keep = New-Object System.Collections.Generic.List[string]
    foreach ($name in @($props.PSObject.Properties.Name | Where-Object { $_ -match '^OPEN\d*$' } | Sort-Object { if ($_ -eq 'OPEN') { 0 } else { [int]($_ -replace '\D','') } })) {
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

    $keep.Add("`"$XlamPath`"")
    for ($i = 0; $i -lt $keep.Count; $i++) {
        $name = if ($i -eq 0) { 'OPEN' } else { "OPEN$i" }
        Set-ItemProperty -Path $key -Name $name -Value $keep[$i]
    }

    Write-Host "Registered Excel add-in auto-open: $XlamPath" -ForegroundColor Green
}

if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

Write-Host "Creating EXLerateShortcuts.xlam..." -ForegroundColor Cyan

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
    $module.Name = 'EXLerateShortcuts'

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
