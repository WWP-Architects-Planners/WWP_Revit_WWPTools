param(
    [string]$BaseRef,
    [string]$HeadRef = "HEAD"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Invoke-Git {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $output = & git @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw ("git {0} failed with exit code {1}" -f ($Arguments -join " "), $LASTEXITCODE)
    }

    return @($output)
}

function Get-ChangedFiles {
    param(
        [string]$BaseRef,
        [string]$HeadRef
    )

    $nullSha = "0000000000000000000000000000000000000000"
    if ([string]::IsNullOrWhiteSpace($BaseRef) -or $BaseRef -eq $nullSha) {
        return Invoke-Git -Arguments @("diff-tree", "--no-commit-id", "--name-only", "-r", $HeadRef)
    }

    return Invoke-Git -Arguments @("diff", "--name-only", $BaseRef, $HeadRef)
}

$changedFiles = @(Get-ChangedFiles -BaseRef $BaseRef -HeadRef $HeadRef |
    ForEach-Object { ($_ -replace "\\", "/").Trim() } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

if (-not $changedFiles) {
    Write-Host "No changed files detected."
    exit 0
}

$wpfUiChanged = @($changedFiles | Where-Object { $_ -like "WWPTools.WpfUI/*" })
$requiredDlls = @(
    "WWPTools.extension/lib/WWPTools.WpfUI.net48.dll",
    "WWPTools.extension/lib/WWPTools.WpfUI.net8.0-windows.dll"
)
$changedDlls = @($changedFiles | Where-Object { $requiredDlls -contains $_ })
$missingDlls = @($requiredDlls | Where-Object { $changedDlls -notcontains $_ })

if ($wpfUiChanged.Count -gt 0 -and $missingDlls.Count -gt 0) {
    Write-Error @"
WWPTools.WpfUI source files changed, but the shipped DLLs were not fully updated.

Changed source files:
$($wpfUiChanged -join [Environment]::NewLine)

Missing DLL updates:
$($missingDlls -join [Environment]::NewLine)

Build WWPTools.WpfUI locally so the CopyToLib target refreshes the shipped DLLs in WWPTools.extension/lib, then commit those DLL changes and push again.
"@
}

Write-Host "Publish input validation passed."
