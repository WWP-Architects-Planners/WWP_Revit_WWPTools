param(
  [string]$RepoOwner = "jason-WWP",
  [string]$RepoName = "WWPTools",
  [string]$Branch = "main",
  [string]$TargetPackages = "C:\\dynpackages"
)

$ErrorActionPreference = "Stop"

$extensionsRoot = Join-Path $env:APPDATA "pyRevit\\Extensions"
$targetExtension = Join-Path $extensionsRoot "WWPTools.extension"

$repoZipUrl = "https://github.com/$RepoOwner/$RepoName/archive/refs/heads/$Branch.zip"

$tempDir = Join-Path $env:TEMP ("WWPTools_" + [Guid]::NewGuid().ToString("N"))
$zipPath = Join-Path $tempDir "$RepoName-$Branch.zip"
$extractDir = Join-Path $tempDir "extract"

New-Item -ItemType Directory -Path $tempDir | Out-Null
New-Item -ItemType Directory -Path $extractDir | Out-Null

Invoke-WebRequest -Uri $repoZipUrl -OutFile $zipPath -UseBasicParsing
Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force

$repoRoot = Join-Path $extractDir "$RepoName-$Branch"
$extractedExtension = Join-Path $repoRoot "WWPTools.extension"
if (-not (Test-Path $extractedExtension)) {
  throw "Expected WWPTools.extension in the repo zip."
}

New-Item -ItemType Directory -Path $extensionsRoot -Force | Out-Null
if (Test-Path $targetExtension) {
  Remove-Item -Path $targetExtension -Recurse -Force
}

Copy-Item -Path $extractedExtension -Destination $targetExtension -Recurse -Force

$packagesSource = Join-Path $repoRoot "packages"
if (Test-Path $packagesSource) {
  New-Item -ItemType Directory -Path $TargetPackages -Force | Out-Null
  Copy-Item -Path (Join-Path $packagesSource "*") -Destination $TargetPackages -Recurse -Force
}

function Ensure-SettingsFile($settingsPath) {
  if (Test-Path $settingsPath) {
    return $false
  }

  $settingsDir = Split-Path -Parent $settingsPath
  New-Item -ItemType Directory -Path $settingsDir -Force | Out-Null

  $xml = New-Object System.Xml.XmlDocument
  $decl = $xml.CreateXmlDeclaration("1.0", "utf-8", $null)
  $xml.AppendChild($decl) | Out-Null
  $root = $xml.CreateElement("PreferenceSettings")
  $cpf = $xml.CreateElement("CustomPackageFolders")
  $root.AppendChild($cpf) | Out-Null
  $xml.AppendChild($root) | Out-Null
  $xml.Save($settingsPath)
  return $true
}

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
$createdSettings = $false
if (Test-Path $dynamoRoot) {
  Get-ChildItem -Path $dynamoRoot -Directory | ForEach-Object {
    $settingsPath = Join-Path $_.FullName "DynamoSettings.xml"
    if (Ensure-SettingsFile -settingsPath $settingsPath) {
      $createdSettings = $true
    }
    Add-PackagePath -settingsPath $settingsPath -packagePath $TargetPackages
  }
}

Remove-Item -Path $tempDir -Recurse -Force

Write-Host "WWPTools installed to $targetExtension"
if (Test-Path $TargetPackages) {
  Write-Host "Dynamo packages installed to $TargetPackages"
}
if ($createdSettings) {
  $msg = "Dynamo settings were created. Please launch Dynamo once to finish setup."
  Write-Host $msg
  try {
    $shell = New-Object -ComObject WScript.Shell
    $shell.Popup($msg, 10, "WWPTools", 64) | Out-Null
  } catch {
    # Fallback to console only.
  }
}

