param(
  [string]$RepoOwner = "jason-svn",
  [string]$RepoName = "WWPTools",
  [string]$AssetName = "WWPTools-DynamoPackages.zip",
  [string]$TargetRoot = "C:\\dynpackages"
)

$ErrorActionPreference = "Stop"

$apiUrl = "https://api.github.com/repos/$RepoOwner/$RepoName/releases/latest"
$release = Invoke-RestMethod -Uri $apiUrl -Headers @{ "User-Agent" = "WWPTools-Installer" }

$asset = $release.assets | Where-Object { $_.name -eq $AssetName } | Select-Object -First 1
if (-not $asset) {
  throw "Release asset '$AssetName' not found. Upload it to the latest GitHub Release."
}

New-Item -ItemType Directory -Path $TargetRoot -Force | Out-Null

$tempDir = Join-Path $env:TEMP ("WWPTools_Packages_" + [Guid]::NewGuid().ToString("N"))
$zipPath = Join-Path $tempDir $asset.name
$extractDir = Join-Path $tempDir "extract"

New-Item -ItemType Directory -Path $tempDir | Out-Null
New-Item -ItemType Directory -Path $extractDir | Out-Null

Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zipPath -UseBasicParsing
Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force

Copy-Item -Path (Join-Path $extractDir "*") -Destination $TargetRoot -Recurse -Force

Remove-Item -Path $tempDir -Recurse -Force

function Add-PackagePath($settingsPath, $packagePath) {
  if (-not (Test-Path $settingsPath)) {
    return
  }

  [xml]$xml = Get-Content -Path $settingsPath
  $prefs = $xml.PreferenceSettings
  if (-not $prefs) {
    return
  }

  $cpf = $prefs.CustomPackageFolders
  if (-not $cpf) {
    $cpf = $xml.CreateElement("CustomPackageFolders")
    $prefs.AppendChild($cpf) | Out-Null
  }

  $existing = @($cpf.string)
  if ($existing -notcontains $packagePath) {
    $node = $xml.CreateElement("string")
    $node.InnerText = $packagePath
    $cpf.AppendChild($node) | Out-Null
    $xml.Save($settingsPath)
  }
}

$dynamoRoot = Join-Path $env:APPDATA "Dynamo\\Dynamo Revit"
if (Test-Path $dynamoRoot) {
  Get-ChildItem -Path $dynamoRoot -Directory | ForEach-Object {
    $settingsPath = Join-Path $_.FullName "DynamoSettings.xml"
    Add-PackagePath -settingsPath $settingsPath -packagePath $TargetRoot
  }
}

Write-Host "Dynamo packages installed to $TargetRoot"

