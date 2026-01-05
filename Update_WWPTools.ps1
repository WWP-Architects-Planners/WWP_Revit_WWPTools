param(
  [string]$RepoOwner = "jason-svn",
  [string]$RepoName = "WWPTools"
)

$ErrorActionPreference = "Stop"

$extensionsRoot = Join-Path $env:APPDATA "pyRevit\Extensions"
$targetExtension = Join-Path $extensionsRoot "pyRevitTool.extension"

$apiUrl = "https://api.github.com/repos/$RepoOwner/$RepoName/releases/latest"
$release = Invoke-RestMethod -Uri $apiUrl -Headers @{ "User-Agent" = "WWPTools-Installer" }

if (-not $release.assets -or $release.assets.Count -eq 0) {
  throw "No release assets found. Upload a zip release asset that contains pyRevitTool.extension."
}

$zipAsset = $release.assets | Where-Object { $_.name -match "\.zip$" } | Select-Object -First 1
if (-not $zipAsset) {
  throw "No .zip asset found. Upload a zip release asset that contains pyRevitTool.extension."
}

$tempDir = Join-Path $env:TEMP ("WWPTools_" + [Guid]::NewGuid().ToString("N"))
$zipPath = Join-Path $tempDir $zipAsset.name
$extractDir = Join-Path $tempDir "extract"

New-Item -ItemType Directory -Path $tempDir | Out-Null
New-Item -ItemType Directory -Path $extractDir | Out-Null

Invoke-WebRequest -Uri $zipAsset.browser_download_url -OutFile $zipPath -UseBasicParsing
Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force

$extractedExtension = Join-Path $extractDir "pyRevitTool.extension"
if (-not (Test-Path $extractedExtension)) {
  throw "Expected pyRevitTool.extension at the root of the zip."
}

New-Item -ItemType Directory -Path $extensionsRoot -Force | Out-Null
if (Test-Path $targetExtension) {
  Remove-Item -Path $targetExtension -Recurse -Force
}

Copy-Item -Path $extractedExtension -Destination $targetExtension -Recurse -Force

Remove-Item -Path $tempDir -Recurse -Force

Write-Host "WWPTools installed to $targetExtension"
