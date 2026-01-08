$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$wxsPath = Join-Path $scriptDir "WWPTools.wxs"

# Extract version from WXS file
[xml]$wxsContent = Get-Content $wxsPath
$version = $wxsContent.Wix.Package.Version
if ($version -match '^(\d+\.\d+\.\d+)') {
    $versionTag = "v$($matches[1])"
    $msiPath = Join-Path $scriptDir "WWPTools-$versionTag.msi"
} else {
    $msiPath = Join-Path $scriptDir "WWPTools.msi"
}

$wixExe = "C:\Program Files\WiX Toolset v6.0\bin\wix.exe"
$bindPath = $scriptDir

if (-not (Test-Path $wixExe)) {
  throw "WiX not found at $wixExe"
}

& $wixExe build -arch x64 -o $msiPath $wxsPath -bindpath $bindPath -ext WixToolset.Util.wixext -ext WixToolset.UI.wixext

$installer = New-Object -ComObject WindowsInstaller.Installer
$db = $installer.OpenDatabase($msiPath, 1)
$view = $db.OpenView("UPDATE `Control` SET `Width`=370 WHERE `Control`='BannerLine' OR `Control`='BottomLine'")
$view.Execute()
$db.Commit()
$view.Close()

Write-Host "MSI built and UI controls patched: $msiPath"
