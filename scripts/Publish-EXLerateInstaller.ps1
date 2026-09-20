<#
.SYNOPSIS
  Builds a per-user office installer: folder payload, ClickOnce publish, optional MSI.
.DESCRIPTION
  Developer machine (VS + Excel) produces artifacts under dist\ that an office laptop
  can install without administrator rights:

    dist\EXLerate-Installer\Install.cmd   <- copy this folder, double-click Install.cmd
    dist\EXLerateSetup.msi                <- if WiX is available
    dist\clickonce\setup.exe              <- VSTO ClickOnce (needs VSTO Runtime)

  Office laptop: copy the installer, close Excel, run it, reopen Excel -> EXLerate tab.
#>
param(
    [switch]$SkipBuild,
    [switch]$SkipXlam,
    [switch]$SkipClickOnce,
    [switch]$SkipMsi,
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    [string]$Version = '1.0.0'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$project = Join-Path $repoRoot 'src\ArixcelExplorer\ArixcelExplorer.csproj'
$binDir = Join-Path $repoRoot "src\ArixcelExplorer\bin\$Configuration"
$distDir = Join-Path $repoRoot 'dist'
$stageDir = Join-Path $distDir 'EXLerate-Installer'
$clickOnceDir = Join-Path $distDir 'clickonce'
$msiPath = Join-Path $distDir 'EXLerateSetup.msi'
$productVersion = if ($Version -match '^\d+\.\d+\.\d+\.\d+$') { $Version } else { "$Version.0" }

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

function ConvertTo-WixId([string]$value) {
    $id = ($value -replace '[^A-Za-z0-9]', '_')
    if ($id -match '^[0-9]') { $id = "F_$id" }
    if ($id.Length -gt 70) { $id = $id.Substring(0, 70) }
    return $id
}

function New-WixPayloadFragment {
    param([string]$StagePath, [string]$OutPath)

    $files = Get-ChildItem -Path $StagePath -Recurse -File
    if (-not $files) { throw "No files staged in $StagePath" }

    $dirs = New-Object System.Collections.Generic.HashSet[string]
    $compRefs = New-Object System.Collections.Generic.List[string]
    $components = New-Object System.Collections.Generic.List[string]

    foreach ($file in $files) {
        $rel = $file.FullName.Substring($StagePath.Length).TrimStart('\')
        $parentRel = Split-Path $rel -Parent
        $dirId = 'INSTALLFOLDER'
        if ($parentRel) {
            $dirId = 'dir_' + (ConvertTo-WixId ($parentRel -replace '\\', '_'))
            [void]$dirs.Add("$dirId|$parentRel")
        }
        $compId = 'cmp_' + (ConvertTo-WixId ($rel -replace '\\', '_'))
        $src = $file.FullName
        $components.Add(@"
      <Component Id="$compId" Directory="$dirId" Guid="*">
        <File Source="$src" KeyPath="yes" />
      </Component>
"@)
        $compRefs.Add("      <ComponentRef Id=`"$compId`" />")
    }

    $dirXml = New-Object System.Collections.Generic.List[string]
    foreach ($entry in ($dirs | Sort-Object)) {
        $parts = $entry.Split('|', 2)
        $name = Split-Path $parts[1] -Leaf
        $dirXml.Add("    <DirectoryRef Id=`"INSTALLFOLDER`">`n      <Directory Id=`"$($parts[0])`" Name=`"$name`" />`n    </DirectoryRef>")
    }

    $xml = @"
<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs">
  <Fragment>
$($dirXml -join "`n")
    <ComponentGroup Id="PayloadComponents">
$($components -join "`n")
$($compRefs -join "`n")
    </ComponentGroup>
  </Fragment>
</Wix>
"@
    # ComponentGroup cannot contain both Component and ComponentRef of the same items.
    $xml = @"
<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs">
  <Fragment>
$($dirXml -join "`n")
    <ComponentGroup Id="PayloadComponents">
$($components -join "`n")
    </ComponentGroup>
  </Fragment>
</Wix>
"@
    Set-Content -Path $OutPath -Value $xml -Encoding UTF8
}

Write-Host ""
Write-Host "EXLerate per-user installer publish" -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor Cyan

if (-not $SkipBuild) {
    & (Join-Path $PSScriptRoot 'Build-ArixcelVsto.ps1') -Configuration $Configuration | Out-Null
}

$vsto = Join-Path $binDir 'ArixcelExplorer.vsto'
$dll = Join-Path $binDir 'ArixcelExplorer.dll'
if (-not (Test-Path $dll) -or -not (Test-Path $vsto)) {
    throw "Build output missing under $binDir. Run without -SkipBuild."
}

$xlamPath = Join-Path $distDir 'EXLerateShortcuts.xlam'
if (-not $SkipXlam) {
    try {
        & (Join-Path $PSScriptRoot 'New-ArixcelShortcutsXlam.ps1') -OutputPath $xlamPath | Out-Null
    } catch {
        Write-Host "Could not build .xlam (Excel busy or VBA access disabled): $($_.Exception.Message)" -ForegroundColor Yellow
        if (-not (Test-Path $xlamPath)) {
            Write-Host "Installer will ship without Ctrl+Q until EXLerateShortcuts.xlam is present." -ForegroundColor Yellow
        }
    }
}

if (Test-Path $stageDir) { Remove-Item $stageDir -Recurse -Force }
New-Item -ItemType Directory -Path $stageDir -Force | Out-Null

$copyExt = @('.dll', '.vsto', '.manifest', '.xml')
Get-ChildItem -Path $binDir -File | Where-Object { $copyExt -contains $_.Extension.ToLowerInvariant() } | ForEach-Object {
    Copy-Item $_.FullName -Destination (Join-Path $stageDir $_.Name) -Force
}
$ribbonSrc = Join-Path $binDir 'Ribbon'
if (Test-Path $ribbonSrc) {
    Copy-Item $ribbonSrc (Join-Path $stageDir 'Ribbon') -Recurse -Force
}

foreach ($name in @(
    'Install-EXLeratePerUser.ps1',
    'Uninstall-EXLeratePerUser.ps1',
    'Register-ArixcelComAddIn.ps1'
)) {
    Copy-Item (Join-Path $PSScriptRoot $name) (Join-Path $stageDir $name) -Force
}
Copy-Item (Join-Path $repoRoot 'installer\Install.cmd') (Join-Path $stageDir 'Install.cmd') -Force
Copy-Item (Join-Path $repoRoot 'installer\Uninstall.cmd') (Join-Path $stageDir 'Uninstall.cmd') -Force
Copy-Item (Join-Path $repoRoot 'installer\OFFICE-INSTALL.txt') (Join-Path $stageDir 'OFFICE-INSTALL.txt') -Force
Copy-Item (Join-Path $repoRoot 'installer\RegisterFromMsi.cmd') (Join-Path $stageDir 'RegisterFromMsi.cmd') -Force
Copy-Item (Join-Path $repoRoot 'installer\UnregisterFromMsi.cmd') (Join-Path $stageDir 'UnregisterFromMsi.cmd') -Force
if (Test-Path $xlamPath) {
    Copy-Item $xlamPath (Join-Path $stageDir 'EXLerateShortcuts.xlam') -Force
}

Write-Host "Staged folder installer: $stageDir" -ForegroundColor Green

if (-not $SkipClickOnce) {
    $msbuild = Get-VsMsBuildPath
    if (-not $msbuild) {
        Write-Host "MSBuild not found; skipping ClickOnce publish." -ForegroundColor Yellow
    } else {
        if (Test-Path $clickOnceDir) { Remove-Item $clickOnceDir -Recurse -Force }
        New-Item -ItemType Directory -Path $clickOnceDir -Force | Out-Null
        $clickOnceBuild = Join-Path $distDir 'clickonce-build'
        if (Test-Path $clickOnceBuild) { Remove-Item $clickOnceBuild -Recurse -Force }
        New-Item -ItemType Directory -Path $clickOnceBuild -Force | Out-Null
        $pfxPath = Join-Path $repoRoot 'src\ArixcelExplorer\ArixcelExplorer_TemporaryKey.pfx'
        $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -like '*Arixcel Explorer Dev*' } | Select-Object -First 1
        $publishArgs = @(
            $project,
            '/t:Publish',
            "/p:Configuration=$Configuration",
            '/p:Platform=AnyCPU',
            "/p:OutputPath=$clickOnceBuild/",
            "/p:ApplicationVersion=$productVersion",
            "/p:PublishUrl=$clickOnceDir/",
            "/p:PublishDir=$clickOnceDir/",
            '/p:InstallFrom=Disk',
            '/p:IsWebBootstrapper=false',
            '/p:BootstrapperEnabled=true',
            '/p:UpdateEnabled=false',
            '/p:SignManifests=true',
            '/p:MSBuildEnableWorkloadResolver=false',
            '/verbosity:minimal'
        )
        if ($cert) {
            $publishArgs += "/p:ManifestCertificateThumbprint=$($cert.Thumbprint)"
        } elseif (Test-Path $pfxPath) {
            $publishArgs += "/p:ManifestKeyFile=$pfxPath"
            $publishArgs += '/p:ManifestCertificatePassword=arixcel-dev'
        }
        Write-Host "Publishing ClickOnce to $clickOnceDir ..." -ForegroundColor Cyan
        & $msbuild @publishArgs
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ClickOnce publish failed (folder installer is still available)." -ForegroundColor Yellow
        } else {
            Write-Host "ClickOnce: $clickOnceDir" -ForegroundColor Green
        }
    }
}

if (-not $SkipMsi) {
    $wix = Get-Command wix -ErrorAction SilentlyContinue
    if (-not $wix) {
        Write-Host "WiX CLI not found. Installing as a local dotnet tool..." -ForegroundColor Yellow
        $manifestDir = Join-Path $repoRoot '.config'
        New-Item -ItemType Directory -Path $manifestDir -Force | Out-Null
        $manifest = Join-Path $manifestDir 'dotnet-tools.json'
        Push-Location $repoRoot
        try {
            if (-not (Test-Path $manifest)) {
                dotnet new tool-manifest --force | Out-Null
            }
            dotnet tool update wix --version 5.0.2 2>$null
            if ($LASTEXITCODE -ne 0) {
                dotnet tool install wix --version 5.0.2
            }
        } finally {
            Pop-Location
        }
        $wixCmd = Join-Path $repoRoot '.config'
        $wix = Get-Command wix -ErrorAction SilentlyContinue
    }

    $wixExe = $null
    if ($wix) {
        $wixExe = $wix.Source
    } else {
        $localWix = Get-ChildItem -Path (Join-Path $repoRoot '.tools') -Filter 'wix.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($localWix) { $wixExe = $localWix.FullName }
    }

    if (-not $wixExe) {
        Push-Location $repoRoot
        try {
            $toolPath = dotnet tool run wix -- --version 2>$null
            if ($LASTEXITCODE -eq 0) { $wixExe = 'dotnet-tool-wix' }
        } catch {
            $wixExe = $null
        } finally {
            Pop-Location
        }
    }

    if (-not $wix -and -not $wixExe) {
        Write-Host "WiX is not available; skipping MSI. Use dist\EXLerate-Installer\Install.cmd on the office laptop." -ForegroundColor Yellow
    } else {
        $payloadWxs = Join-Path $repoRoot 'installer\Payload.g.wxs'
        New-WixPayloadFragment -StagePath $stageDir -OutPath $payloadWxs
        if (Test-Path $msiPath) { Remove-Item $msiPath -Force }
        Write-Host "Building per-user MSI..." -ForegroundColor Cyan
        Push-Location $repoRoot
        try {
            if ($wix) {
                & $wix.Source build `
                    (Join-Path $repoRoot 'installer\EXLerate.wxs') `
                    $payloadWxs `
                    -d "ProductVersion=$productVersion" `
                    -arch x64 `
                    -o $msiPath
            } else {
                dotnet tool run wix -- build `
                    (Join-Path $repoRoot 'installer\EXLerate.wxs') `
                    $payloadWxs `
                    -d "ProductVersion=$productVersion" `
                    -arch x64 `
                    -o $msiPath
            }
        } finally {
            Pop-Location
        }
        if (Test-Path $msiPath) {
            Write-Host "MSI: $msiPath" -ForegroundColor Green
        } else {
            Write-Host "MSI build did not produce $msiPath. Folder installer is still available." -ForegroundColor Yellow
        }
    }
}

Write-Host @"

  Publish complete
  ================
  Office laptop (no admin):

    1. Copy dist\EXLerateSetup.msi  OR  the folder dist\EXLerate-Installer
    2. Close Excel
    3. Double-click the MSI, or run Install.cmd
    4. Reopen Excel - the EXLerate tab should appear

  Folder: $stageDir
  MSI:    $(if (Test-Path $msiPath) { $msiPath } else { '(not built)' })
  ClickOnce: $(if (Test-Path (Join-Path $clickOnceDir 'setup.exe')) { Join-Path $clickOnceDir 'setup.exe' } else { '(not built)' })

"@ -ForegroundColor Green

