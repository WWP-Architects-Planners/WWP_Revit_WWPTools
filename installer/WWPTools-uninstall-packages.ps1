$ErrorActionPreference = 'Stop'

$revitProcs = Get-Process -Name 'revit' -ErrorAction SilentlyContinue
if ($revitProcs) {
    $msg = 'Please close Revit before uninstalling WWPTools packages. The uninstaller cannot update files while Revit is running.'
    try {
        $shell = New-Object -ComObject WScript.Shell
        $shell.Popup($msg, 0, 'WWPTools Installer', 48) | Out-Null
    } catch {
        Write-Host $msg
    }
    throw 'Revit is running, uninstall aborted.'
}

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
