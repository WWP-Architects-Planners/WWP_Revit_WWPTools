$ErrorActionPreference = 'Stop'

$revitProcs = Get-Process -Name 'revit' -ErrorAction SilentlyContinue
if ($revitProcs) {
    $msg = 'Please close Revit before uninstalling WWPTools. The uninstaller cannot update files while Revit is running.'
    try {
        $shell = New-Object -ComObject WScript.Shell
        $shell.Popup($msg, 0, 'WWPTools Installer', 48) | Out-Null
    } catch {
        Write-Host $msg
    }
    throw 'Revit is running, uninstall aborted.'
}

$targetExtension = Join-Path $env:APPDATA 'pyRevit\Extensions\WWPTools.extension'
if (Test-Path $targetExtension) {
    Remove-Item -Path $targetExtension -Recurse -Force -ErrorAction SilentlyContinue
}
