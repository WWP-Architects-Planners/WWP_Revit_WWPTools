$ErrorActionPreference = 'Stop'

$targetExtension = Join-Path $env:APPDATA 'pyRevit\Extensions\WWPTools.extension'
if (Test-Path $targetExtension) {
    Remove-Item -Path $targetExtension -Recurse -Force -ErrorAction SilentlyContinue
}
