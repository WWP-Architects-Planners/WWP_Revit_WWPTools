Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-PackageVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WxsPath
    )

    [xml]$wxs = Get-Content -LiteralPath $WxsPath
    $package = $wxs.Wix.Package
    if (-not $package) {
        throw "Could not find <Package> in $WxsPath."
    }

    $version = $package.Version
    if (-not $version) {
        throw "Package Version is missing in $WxsPath."
    }

    if ($version -match "^\d+\.\d+\.\d+\.0$") {
        return $version.Substring(0, $version.Length - 2)
    }

    return $version
}

function Ensure-Wix {
    $wix = Get-Command wix -ErrorAction SilentlyContinue
    if (-not $wix) {
        if ($env:WIX_BIN) {
            $wixPath = $env:WIX_BIN
            if ((Test-Path -LiteralPath $wixPath) -and (Get-Item $wixPath).PSIsContainer) {
                $wixPath = Join-Path $wixPath "wix.exe"
            }
            if (Test-Path -LiteralPath $wixPath) {
                return $wixPath
            }
        }

        $candidates = @(
            (Join-Path $env:ProgramFiles "WiX Toolset v4\\bin\\wix.exe"),
            (Join-Path $env:ProgramFiles "WiX Toolset v5\\bin\\wix.exe"),
            (Join-Path $env:ProgramFiles "WiX Toolset v6.0\\bin\\wix.exe"),
            (Join-Path $env:ProgramFiles "WiX Toolset v4\\wix.exe"),
            (Join-Path $env:ProgramFiles "WiX Toolset v5\\wix.exe"),
            (Join-Path $env:ProgramFiles "WiX Toolset v6.0\\wix.exe")
        )

        $programFilesX86 = ${env:ProgramFiles(x86)}
        if ($programFilesX86) {
            $candidates += Join-Path $programFilesX86 "WiX Toolset v4\\bin\\wix.exe"
            $candidates += Join-Path $programFilesX86 "WiX Toolset v5\\bin\\wix.exe"
            $candidates += Join-Path $programFilesX86 "WiX Toolset v6.0\\bin\\wix.exe"
            $candidates += Join-Path $programFilesX86 "WiX Toolset v4\\wix.exe"
            $candidates += Join-Path $programFilesX86 "WiX Toolset v5\\wix.exe"
            $candidates += Join-Path $programFilesX86 "WiX Toolset v6.0\\wix.exe"
        }

        foreach ($candidate in $candidates) {
            if ($candidate -and (Test-Path -LiteralPath $candidate)) {
                return $candidate
            }
        }

        throw "WiX CLI (wix.exe) not found. Install WiX Toolset v4+ and ensure wix.exe is on PATH or set WIX_BIN to its folder."
    }

    return $wix.Path
}

function Patch-MsiDialogMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MsiPath
    )

    $installer = New-Object -ComObject WindowsInstaller.Installer
    $db = $installer.OpenDatabase($MsiPath, 1)
    $fixes = 0

    $dialogView = $db.OpenView('SELECT `Dialog`, `Control_First` FROM `Dialog`')
    $dialogView.Execute()

    while ($dialogRecord = $dialogView.Fetch()) {
        $dialogId = $dialogRecord.StringData(1)
        $controlFirst = $dialogRecord.StringData(2)
        if ([string]::IsNullOrWhiteSpace($controlFirst)) {
            continue
        }

        $controlQuery = 'SELECT `Control` FROM `Control` WHERE `Dialog_`=''{0}'' AND `Control`=''{1}''' -f $dialogId, $controlFirst
        $controlView = $db.OpenView($controlQuery)
        $controlView.Execute()
        $controlRecord = $controlView.Fetch()

        if (-not $controlRecord) {
            $firstQuery = 'SELECT `Control` FROM `Control` WHERE `Dialog_`=''{0}'' ORDER BY `TabOrder`' -f $dialogId
            $firstView = $db.OpenView($firstQuery)
            $firstView.Execute()
            $firstRecord = $firstView.Fetch()
            if ($firstRecord) {
                $replacement = $firstRecord.StringData(1)
                $updateView = $db.OpenView('UPDATE `Dialog` SET `Control_First`=? WHERE `Dialog`=?')
                $updateRecord = $installer.CreateRecord(2)
                $updateRecord.StringData(1) = $replacement
                $updateRecord.StringData(2) = $dialogId
                $updateView.Execute($updateRecord)
                $fixes++
            }
        }
    }

    if ($fixes -gt 0) {
        $db.Commit()
    }

    return $fixes
}

function Get-EncodedCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath
    )

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "Script file not found at $ScriptPath."
    }

    $scriptText = Get-Content -LiteralPath $ScriptPath -Raw
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($scriptText)
    return [Convert]::ToBase64String($bytes)
}

function Build-WwpToolsMsi {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WxsPath,
        [Parameter(Mandatory = $true)]
        [string]$InstallScriptPath,
        [Parameter(Mandatory = $true)]
        [string]$UninstallScriptPath,
        [Parameter(Mandatory = $true)]
        [string]$OutputPrefix
    )

    if (-not (Test-Path -LiteralPath $WxsPath)) {
        throw "Expected Wix source at $WxsPath."
    }
    if (-not (Test-Path -LiteralPath $InstallScriptPath)) {
        throw "Install script not found at $InstallScriptPath."
    }
    if (-not (Test-Path -LiteralPath $UninstallScriptPath)) {
        throw "Uninstall script not found at $UninstallScriptPath."
    }

    $encodedInstall = Get-EncodedCommand -ScriptPath $InstallScriptPath
    $encodedUninstall = Get-EncodedCommand -ScriptPath $UninstallScriptPath
    $wxsBuildPath = Join-Path $PSScriptRoot ((Split-Path -Leaf $WxsPath) + ".generated")
    $wxsContent = Get-Content -LiteralPath $WxsPath -Raw
    if ($wxsContent -notmatch "__WWPTOOLS_PS_INSTALL_BASE64__") {
        throw "Placeholder __WWPTOOLS_PS_INSTALL_BASE64__ not found in $WxsPath."
    }
    if ($wxsContent -notmatch "__WWPTOOLS_PS_UNINSTALL_BASE64__") {
        throw "Placeholder __WWPTOOLS_PS_UNINSTALL_BASE64__ not found in $WxsPath."
    }
    $wxsContent = $wxsContent.Replace("__WWPTOOLS_PS_INSTALL_BASE64__", $encodedInstall)
    $wxsContent = $wxsContent.Replace("__WWPTOOLS_PS_UNINSTALL_BASE64__", $encodedUninstall)
    Set-Content -LiteralPath $wxsBuildPath -Value $wxsContent -Encoding Ascii

    $version = Get-PackageVersion -WxsPath $WxsPath
    $msiPath = Join-Path $PSScriptRoot ("{0}-v{1}.msi" -f $OutputPrefix, $version)

    if ($OutputPrefix -eq "WWPTools") {
        $versionPath = Join-Path $repoRoot "WWPTools.extension\\lib\\WWPTools.version.json"
        $versionPayload = @{ version = $version } | ConvertTo-Json -Compress
        Set-Content -LiteralPath $versionPath -Value $versionPayload -Encoding Ascii
        Write-Host ("Updated extension version file: {0}" -f $versionPath)

        $aboutBundle = Join-Path $repoRoot "WWPTools.extension\\WWPTools.tab\\6. Links.panel\\About us.urlbutton\\bundle.yaml"
        if (Test-Path -LiteralPath $aboutBundle) {
            $bundleLines = Get-Content -LiteralPath $aboutBundle
            $tooltipLine = "tooltip: \"Installed version: {0}\"" -f $version
            if ($bundleLines -match "^tooltip:") {
                $bundleLines = $bundleLines -replace "^tooltip:.*$", $tooltipLine
            } else {
                $bundleLines = @($bundleLines + $tooltipLine)
            }
            Set-Content -LiteralPath $aboutBundle -Value $bundleLines -Encoding Ascii
            Write-Host ("Updated About tooltip: {0}" -f $aboutBundle)
        }
    }

    Write-Host ("Building {0} installer v{1}..." -f $OutputPrefix, $version)
    Push-Location $PSScriptRoot
    try {
        & $wixExe build $wxsBuildPath -ext WixToolset.UI.wixext -ext WixToolset.Util.wixext -bindpath $PSScriptRoot -o $msiPath
        if ($LASTEXITCODE -ne 0) {
            throw "WiX build failed with exit code $LASTEXITCODE."
        }
    } finally {
        Pop-Location
        if (Test-Path -LiteralPath $wxsBuildPath) {
            Remove-Item -LiteralPath $wxsBuildPath -Force -ErrorAction SilentlyContinue
        }
    }

    $fixCount = Patch-MsiDialogMetadata -MsiPath $msiPath
    if ($fixCount -gt 0) {
        Write-Host ("Patched dialog metadata on {0} dialog(s)." -f $fixCount)
    }

    if (-not (Test-Path -LiteralPath $msiPath)) {
        throw "Build completed but MSI was not found at $msiPath."
    }

    Write-Host ("MSI ready: {0}" -f $msiPath)
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$wixExe = Ensure-Wix

Build-WwpToolsMsi `
    -WxsPath (Join-Path $PSScriptRoot "WWPTools.wxs") `
    -InstallScriptPath (Join-Path $PSScriptRoot "WWPTools-install-scripts.ps1") `
    -UninstallScriptPath (Join-Path $PSScriptRoot "WWPTools-uninstall-scripts.ps1") `
    -OutputPrefix "WWPTools"

Build-WwpToolsMsi `
    -WxsPath (Join-Path $PSScriptRoot "WWPTools-Packages.wxs") `
    -InstallScriptPath (Join-Path $PSScriptRoot "WWPTools-install-packages.ps1") `
    -UninstallScriptPath (Join-Path $PSScriptRoot "WWPTools-uninstall-packages.ps1") `
    -OutputPrefix "WWPTools-Packages"
