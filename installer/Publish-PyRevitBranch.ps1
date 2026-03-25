param(
    [string]$SourceRef = "HEAD",
    [string]$TargetBranch = "pyrevit",
    [string]$ExtensionPath = "WWPTools.extension",
    [string]$Remote = "origin",
    [switch]$NoPush
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Invoke-Git {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    & git @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw ("git {0} failed with exit code {1}" -f ($Arguments -join " "), $LASTEXITCODE)
    }
}

if (-not (Test-Path -LiteralPath $ExtensionPath -PathType Container)) {
    throw ("Extension path not found: {0}" -f $ExtensionPath)
}

$repoRoot = & git rev-parse --show-toplevel
if ($LASTEXITCODE -ne 0) {
    throw "This script must run inside a git repository."
}

Set-Location $repoRoot

Invoke-Git -Arguments @("config", "user.name", "github-actions[bot]")
Invoke-Git -Arguments @("config", "user.email", "41898282+github-actions[bot]@users.noreply.github.com")

$splitCommit = (& git subtree split --prefix $ExtensionPath $SourceRef).Trim()
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($splitCommit)) {
    throw ("Unable to create subtree split for {0} from {1}" -f $ExtensionPath, $SourceRef)
}

if ($NoPush) {
    Write-Host ("Created subtree split commit {0} for {1}. Skipping push because -NoPush was specified." -f $splitCommit, $ExtensionPath)
    exit 0
}

Write-Host ("Publishing {0} from {1} to {2}/{3}" -f $ExtensionPath, $SourceRef, $Remote, $TargetBranch)
Invoke-Git -Arguments @("push", $Remote, ("{0}:refs/heads/{1}" -f $splitCommit, $TargetBranch), "--force")
Write-Host ("Published commit {0} to branch {1}" -f $splitCommit, $TargetBranch)
