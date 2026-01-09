# WWPTools Release Publishing Script
param(
    [string]$Version = "1.1.2",
    [string]$Message = "Update changelog for v$Version release"
)

$WorkspaceRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $WorkspaceRoot

Write-Host "=== WWPTools Release Publisher ===" -ForegroundColor Cyan
Write-Host "Version: $Version" -ForegroundColor Yellow

# Git config
Write-Host "[1/4] Configuring Git..." -ForegroundColor Cyan
git config user.name jason-svn
Write-Host "OK" -ForegroundColor Green

# Commit and push
Write-Host "[2/4] Committing and pushing..." -ForegroundColor Cyan
git add CHANGELOG.md
git commit -m "chore: $Message"
git push origin main
Write-Host "OK" -ForegroundColor Green

# Check installer
Write-Host "[3/4] Verifying installer..." -ForegroundColor Cyan
$installer = "installer/WWPTools-v$Version.msi"
if (-not (Test-Path $installer)) {
    Write-Host "Installer not found: $installer" -ForegroundColor Red
    exit 1
}
Write-Host "OK" -ForegroundColor Green

# Upload to release
Write-Host "[4/4] Uploading to GitHub release..." -ForegroundColor Cyan
gh release upload "v$Version" $installer --clobber
Write-Host "OK" -ForegroundColor Green

Write-Host ""
Write-Host "Release published successfully!" -ForegroundColor Green
Write-Host "https://github.com/jason-svn/WWPTools/releases/tag/v$Version" -ForegroundColor Cyan
