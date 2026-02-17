$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$logPath = Join-Path $env:TEMP 'WWPTools-install.log'
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    try {
        Add-Content -Path $logPath -Value ("{0} {1}" -f $timestamp, $Message) -Encoding Ascii
    } catch {
    }
}

function Download-File {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [Parameter(Mandatory = $true)]
        [string]$OutFile,
        [int]$TimeoutSeconds = 300
    )

    $client = New-Object System.Net.WebClient
    $client.Proxy = [Net.WebRequest]::GetSystemWebProxy()
    $client.Proxy.Credentials = [Net.CredentialCache]::DefaultCredentials
    try {
        Write-Log ("Downloading {0} -> {1}" -f $Uri, $OutFile)
        $task = $client.DownloadFileTaskAsync($Uri, $OutFile)
        if (-not $task.Wait([TimeSpan]::FromSeconds($TimeoutSeconds))) {
            throw "Download timed out after $TimeoutSeconds seconds."
        }
        Write-Log "Download completed."
    } catch {
        throw "Failed to download $Uri. $($_.Exception.Message)"
    } finally {
        $client.Dispose()
    }
}

Write-Log "=== WWPTools install start ==="
try {
    $repoOwner = 'jason-svn'
    $repoName = 'WWPTools'
    $branch = 'main'
    $targetPackages = 'C:\dynpackages'
    $targetExtension = Join-Path $env:APPDATA 'pyRevit\Extensions\WWPTools.extension'
    $extensionsRoot = Split-Path -Parent $targetExtension
    $repoZipUrl = 'https://github.com/' + $repoOwner + '/' + $repoName + '/archive/refs/heads/' + $branch + '.zip'

    $tempDir = Join-Path $env:TEMP ('WWPTools_' + (New-Guid).ToString('N'))
    $zipPath = Join-Path $tempDir ($repoName + '-' + $branch + '.zip')
    $extractDir = Join-Path $tempDir 'extract'

    Write-Log ("Temp dir: {0}" -f $tempDir)
    New-Item -ItemType Directory -Path $tempDir | Out-Null
    New-Item -ItemType Directory -Path $extractDir | Out-Null
    Download-File -Uri $repoZipUrl -OutFile $zipPath

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    try {
        Write-Log ("Extracting {0} -> {1}" -f $zipPath, $extractDir)
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extractDir)
    } catch {
        throw $_
    }

    $repoRootItem = Get-ChildItem -Path $extractDir -Directory | Select-Object -First 1
    if (-not $repoRootItem) {
        throw 'Expected repo folder in extracted zip.'
    }
    $repoRoot = $repoRootItem.FullName
    Write-Log ("Repo root: {0}" -f $repoRoot)
    $extractedExtension = Join-Path $repoRoot 'WWPTools.extension'
    if (-not (Test-Path $extractedExtension)) {
        throw 'Expected WWPTools.extension in the repo zip.'
    }

    New-Item -ItemType Directory -Path $extensionsRoot -Force | Out-Null
    if (Test-Path $targetExtension) {
        Write-Log ("Removing existing extension at {0}" -f $targetExtension)
        try {
            Remove-Item -Path $targetExtension -Recurse -Force -ErrorAction Stop
        } catch {
            Write-Log ("Failed to remove {0}; clearing contents instead." -f $targetExtension)
            Get-ChildItem -Path $targetExtension -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    New-Item -ItemType Directory -Path $targetExtension -Force | Out-Null
    Write-Log ("Copying extension contents to {0}" -f $targetExtension)
    Copy-Item -Path (Join-Path $extractedExtension '*') -Destination $targetExtension -Recurse -Force

    $packagesSource = Join-Path $repoRoot 'packages'
    if (Test-Path $packagesSource) {
        Write-Log ("Copying packages to {0}" -f $targetPackages)
        New-Item -ItemType Directory -Path $targetPackages -Force | Out-Null
        Copy-Item -Path (Join-Path $packagesSource '*') -Destination $targetPackages -Recurse -Force
    }

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

    try {
        New-Item -Path 'HKCU:\Software\WWP\WWPTools' -Force | Out-Null
        New-ItemProperty -Path 'HKCU:\Software\WWP\WWPTools' -Name 'PackagesInstalledByMain' -PropertyType DWord -Value 1 -Force | Out-Null
    } catch {
        Write-Log ("WARN: failed to set PackagesInstalledByMain marker. {0}" -f $_.Exception.Message)
    }

    Write-Log "Install completed."
} catch {
    Write-Log ("ERROR: {0}" -f $_.Exception.Message)
    if ($_.Exception.InnerException) {
        Write-Log ("INNER: {0}" -f $_.Exception.InnerException.Message)
    }
    if ($_.ScriptStackTrace) {
        Write-Log ("STACK: {0}" -f $_.ScriptStackTrace)
    }
    throw
} finally {
    Write-Log "=== WWPTools install end ==="
}
