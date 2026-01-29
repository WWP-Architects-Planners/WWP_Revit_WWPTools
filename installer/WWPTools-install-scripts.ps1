$ErrorActionPreference = 'Stop'

$revitProcs = Get-Process -Name 'revit' -ErrorAction SilentlyContinue
if ($revitProcs) {
    $msg = 'Please close Revit before installing WWPTools. The installer cannot update files while Revit is running.'
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
$targetExtension = Join-Path $env:APPDATA 'pyRevit\Extensions\WWPTools.extension'
$extensionsRoot = Split-Path -Parent $targetExtension
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
$extractedExtension = Join-Path $repoRoot 'WWPTools.extension'
if (-not (Test-Path $extractedExtension)) {
    throw 'Expected WWPTools.extension in the repo zip.'
}

New-Item -ItemType Directory -Path $extensionsRoot -Force | Out-Null
if (Test-Path $targetExtension) {
    Remove-Item -Path $targetExtension -Recurse -Force -ErrorAction SilentlyContinue
}
Copy-Item -Path $extractedExtension -Destination $targetExtension -Recurse -Force

if (Test-Path $tempDir) {
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}
