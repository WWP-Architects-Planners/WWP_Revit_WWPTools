$ErrorActionPreference = 'Stop'

$revitProcs = Get-Process -Name 'revit' -ErrorAction SilentlyContinue
if ($revitProcs) {
    $msg = 'Please close Revit before installing WWPTools packages. The installer cannot update files while Revit is running.'
    try {
        $shell = New-Object -ComObject WScript.Shell
        $shell.Popup($msg, 0, 'WWPTools Installer', 48) | Out-Null
    } catch {
        Write-Host $msg
    }
    throw 'Revit is running, installation aborted.'
}

$repoOwner = 'jason-svn'
$repoName = 'WWPTools'
$branch = 'main'
$targetPackages = 'C:\dynpackages'
$manifestPath = Join-Path $targetPackages 'WWPTools-packages-manifest.txt'
$repoZipUrl = 'https://github.com/' + $repoOwner + '/' + $repoName + '/archive/refs/heads/' + $branch + '.zip'

$tempDir = Join-Path $env:TEMP ('WWPTools_' + (New-Guid).ToString('N'))
$zipPath = Join-Path $tempDir ($repoName + '-' + $branch + '.zip')
$extractDir = Join-Path $tempDir 'extract'

New-Item -ItemType Directory -Path $tempDir | Out-Null
New-Item -ItemType Directory -Path $extractDir | Out-Null
Invoke-WebRequest -Uri $repoZipUrl -OutFile $zipPath -UseBasicParsing

Add-Type -AssemblyName System.IO.Compression.FileSystem
try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extractDir)
} catch {
    throw $_
}

$repoRoot = Join-Path $extractDir ($repoName + '-' + $branch)
$packagesSource = Join-Path $repoRoot 'packages'
if (-not (Test-Path $packagesSource)) {
    throw 'Expected packages folder in the repo zip.'
}

New-Item -ItemType Directory -Path $targetPackages -Force | Out-Null

if (Test-Path $manifestPath) {
    Get-Content -Path $manifestPath | ForEach-Object {
        $name = $_.Trim()
        if ($name) {
            $installedPath = Join-Path $targetPackages $name
            if (Test-Path $installedPath) {
                Remove-Item -Path $installedPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

$packageNames = @()
Get-ChildItem -Path $packagesSource -Directory | ForEach-Object {
    $packageNames += $_.Name
    Copy-Item -Path $_.FullName -Destination (Join-Path $targetPackages $_.Name) -Recurse -Force
}

$packageNames | Sort-Object | Set-Content -Path $manifestPath -Encoding Ascii

function Ensure-SettingsFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$settingsPath
    )

    if (Test-Path $settingsPath) {
        return $false
    }

    $settingsDir = Split-Path -Parent $settingsPath
    New-Item -ItemType Directory -Path $settingsDir -Force | Out-Null
    $xml = New-Object System.Xml.XmlDocument
    $decl = $xml.CreateXmlDeclaration('1.0', 'utf-8', $null)
    $xml.AppendChild($decl) | Out-Null
    $root = $xml.CreateElement('PreferenceSettings')
    $cpf = $xml.CreateElement('CustomPackageFolders')
    $root.AppendChild($cpf) | Out-Null
    $xml.AppendChild($root) | Out-Null
    $xml.Save($settingsPath)
    return $true
}

function Add-PackagePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$settingsPath,
        [Parameter(Mandatory = $true)]
        [string]$packagePath
    )

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
        $cpf = $xml.CreateElement('CustomPackageFolders')
        $prefs.AppendChild($cpf) | Out-Null
    }

    $existing = @($cpf.string)
    if ($existing -notcontains $packagePath) {
        $node = $xml.CreateElement('string')
        $node.InnerText = $packagePath
        $cpf.AppendChild($node) | Out-Null
        $xml.Save($settingsPath)
    }
}

$dynamoRoot = Join-Path $env:APPDATA 'Dynamo\Dynamo Revit'
$createdSettings = $false
if (Test-Path $dynamoRoot) {
    Get-ChildItem -Path $dynamoRoot -Directory | ForEach-Object {
        $settingsPath = Join-Path $_.FullName 'DynamoSettings.xml'
        if (Ensure-SettingsFile -settingsPath $settingsPath) {
            $createdSettings = $true
        }
        Add-PackagePath -settingsPath $settingsPath -packagePath $targetPackages
    }
}

if (Test-Path $tempDir) {
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

if ($createdSettings) {
    $msg = 'Dynamo settings were created. Please launch Dynamo once to finish setup.'
    try {
        $shell = New-Object -ComObject WScript.Shell
        $shell.Popup($msg, 10, 'WWPTools', 64) | Out-Null
    } catch {
    }
}
