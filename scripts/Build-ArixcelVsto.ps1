<#
.SYNOPSIS
  Builds the ArixcelExplorer VSTO Excel add-in and produces a .vsto manifest.
#>
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$project = Join-Path $repoRoot 'src\ArixcelExplorer\ArixcelExplorer.csproj'
$pfxPath = Join-Path $repoRoot 'src\ArixcelExplorer\ArixcelExplorer_TemporaryKey.pfx'

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

if (-not (Test-Path $pfxPath)) {
    Write-Host "Generating dev signing certificate..." -ForegroundColor Yellow
    $cert = New-SelfSignedCertificate `
        -Subject 'CN=Arixcel Explorer Dev' `
        -Type CodeSigningCert `
        -CertStoreLocation 'Cert:\CurrentUser\My' `
        -KeyExportPolicy Exportable `
        -KeyLength 2048 `
        -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears(5)
    $secure = ConvertTo-SecureString -String 'arixcel-dev' -Force -AsPlainText
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $secure | Out-Null
    Write-Host "Created $pfxPath (password: arixcel-dev)" -ForegroundColor Green
}

$msbuild = Get-VsMsBuildPath
if (-not $msbuild) {
    throw "Visual Studio MSBuild not found. Run .\scripts\Install-ArixcelPrerequisites.ps1 first."
}

Write-Host "Building ArixcelExplorer ($Configuration)..." -ForegroundColor Cyan
& $msbuild $project `
    /restore `
    /p:Configuration=$Configuration `
    /p:Platform=AnyCPU `
    /p:SignManifests=true `
    /p:ManifestKeyFile=$pfxPath `
    /p:ManifestCertificatePassword=arixcel-dev `
    /verbosity:minimal

$outputDir = Join-Path $repoRoot "src\ArixcelExplorer\bin\$Configuration"
$vsto = Join-Path $outputDir 'ArixcelExplorer.vsto'
$dll = Join-Path $outputDir 'ArixcelExplorer.dll'

if (-not (Test-Path $dll)) {
    throw "Build failed - ArixcelExplorer.dll not found in $outputDir"
}
if (-not (Test-Path $vsto)) {
    throw "Build succeeded but .vsto manifest missing. Ensure Office VSTO build tools are installed."
}

Write-Host "Build output:" -ForegroundColor Green
Write-Host "  DLL:  $dll"
Write-Host "  VSTO: $vsto"

return @{
    OutputDir = $outputDir
    DllPath   = $dll
    VstoPath  = $vsto
}
