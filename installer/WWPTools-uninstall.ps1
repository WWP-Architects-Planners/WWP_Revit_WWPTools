$ErrorActionPreference = 'Stop'

$targetExtension = Join-Path $env:APPDATA 'pyRevit\Extensions\WWPTools.extension'
if (Test-Path $targetExtension) {
    Remove-Item -Path $targetExtension -Recurse -Force -ErrorAction SilentlyContinue
}

$markerPath = 'HKCU:\Software\WWP\WWPTools'
$markerName = 'PackagesInstalledByMain'
$installedByMain = $false

try {
    $markerValue = Get-ItemPropertyValue -Path $markerPath -Name $markerName -ErrorAction SilentlyContinue
    if ($markerValue -eq 1) {
        $installedByMain = $true
    }
} catch {
}

if ($installedByMain) {
    $targetPackages = 'C:\dynpackages'
    $manifestPath = Join-Path $targetPackages 'WWPTools-packages-manifest.txt'
    $fallbackNames = @(
        'archi-lab.net',
        'bimorphNodes',
        'Clockwork for Dynamo 2.x',
        'Crumple',
        'Data-Shapes',
        'DynamoIronPython2.7',
        'Elk',
        'Rhythm',
        'spring nodes',
        'Synthetic',
        'WWP Packages'
    )

    $names = @()
    if (Test-Path $manifestPath) {
        $names = Get-Content -Path $manifestPath
    } else {
        $names = $fallbackNames
    }

    foreach ($name in $names) {
        $cleanName = $name.Trim()
        if (-not $cleanName) {
            continue
        }
        $packagePath = Join-Path $targetPackages $cleanName
        if (Test-Path $packagePath) {
            Remove-Item -Path $packagePath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    if (Test-Path $manifestPath) {
        Remove-Item -Path $manifestPath -Force -ErrorAction SilentlyContinue
    }
}

try {
    Remove-ItemProperty -Path $markerPath -Name $markerName -ErrorAction SilentlyContinue
} catch {
}
