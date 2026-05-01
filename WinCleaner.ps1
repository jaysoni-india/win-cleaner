#Requires -Version 5.1
<#
.SYNOPSIS
    WinCleaner — Master Cleanup Orchestrator

.DESCRIPTION
    Interactive master script that runs all three cleanup steps in sequence
    or individually:

      Step 1 — Dev Artifact Cleanup
               Finds and removes dev build/cache folders listed in cleanup-list.txt
               (node_modules, __pycache__, .gradle, .next, dist, etc.)

      Step 2 — Duplicate File Cleanup
               Finds files with identical content (SHA-256 hash) and removes
               duplicate copies, keeping your chosen original.

      Step 3 — Empty File & Folder Cleanup
               Removes zero-byte files and completely empty directory trees.

    All steps generate an HTML report that opens in your browser before any
    deletion occurs. A text deletion log is saved for every run.

.PARAMETER SearchPaths
    One or more root paths to scan. Passed to all steps.
    If omitted, you are prompted once and the same paths apply to all steps.
    Default when prompted with no input: D:\

.PARAMETER Step
    Run a specific step directly without showing the menu.
    Valid values: 1, 2, 3

.PARAMETER DryRun
    Generate reports only; do not delete anything in any step.

.EXAMPLE
    .\WinCleaner.ps1
    .\WinCleaner.ps1 -SearchPaths "D:\Projects","E:\Work"
    .\WinCleaner.ps1 -Step 1
    .\WinCleaner.ps1 -SearchPaths "D:\" -DryRun
#>
[CmdletBinding()]
param(
    [string[]]$SearchPaths = @(),
    [ValidateSet('1','2','3','')]
    [string]  $Step        = '',
    [switch]  $DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# ── Locate step scripts ────────────────────────────────────────────────────────
$ScriptDir = $PSScriptRoot

$Step1Script = Join-Path $ScriptDir 'Step1-DevFolderCleanup.ps1'
$Step2Script = Join-Path $ScriptDir 'Step2-DuplicateFinder.ps1'
$Step3Script = Join-Path $ScriptDir 'Step3-EmptyItemsCleanup.ps1'

foreach ($s in @($Step1Script, $Step2Script, $Step3Script)) {
    if (-not (Test-Path $s)) {
        Write-Host "ERROR: Required script not found: $s" -ForegroundColor Red
        Write-Host 'Ensure all step scripts are in the same folder as WinCleaner.ps1.' -ForegroundColor Yellow
        exit 1
    }
}

# ── Load config.json ───────────────────────────────────────────────────────────
$ConfigFile = Join-Path $ScriptDir 'config.json'
$Config = [PSCustomObject]@{ defaultScanPaths = @('D:\') }
if (Test-Path -LiteralPath $ConfigFile) {
    try {
        $parsed = Get-Content -LiteralPath $ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($null -ne $parsed.defaultScanPaths -and $parsed.defaultScanPaths.Count -gt 0) {
            $Config = $parsed
        }
    }
    catch {
        Write-Host "  WARNING: Could not parse config.json — using defaults." -ForegroundColor Yellow
    }
}

# ── Banner ─────────────────────────────────────────────────────────────────────
function Show-Banner {
    $width = 62
    $line  = '═' * $width
    Write-Host ''
    Write-Host "  $line"  -ForegroundColor Cyan
    Write-Host '  ██╗    ██╗██╗███╗   ██╗ ██████╗██╗     ███████╗ █████╗ ███╗   ██╗███████╗██████╗'  -ForegroundColor Cyan
    Write-Host '  ██║    ██║██║████╗  ██║██╔════╝██║     ██╔════╝██╔══██╗████╗  ██║██╔════╝██╔══██╗' -ForegroundColor Cyan
    Write-Host '  ██║ █╗ ██║██║██╔██╗ ██║██║     ██║     █████╗  ███████║██╔██╗ ██║█████╗  ██████╔╝' -ForegroundColor Cyan
    Write-Host '  ██║███╗██║██║██║╚██╗██║██║     ██║     ██╔══╝  ██╔══██║██║╚██╗██║██╔══╝  ██╔══██╗' -ForegroundColor Cyan
    Write-Host '  ╚███╔███╔╝██║██║ ╚████║╚██████╗███████╗███████╗██║  ██║██║ ╚████║███████╗██║  ██║' -ForegroundColor Cyan
    Write-Host '   ╚══╝╚══╝ ╚═╝╚═╝  ╚═══╝ ╚═════╝╚══════╝╚══════╝╚═╝  ╚═╝╚═╝  ╚═══╝╚══════╝╚═╝  ╚═╝' -ForegroundColor Cyan
    Write-Host "  $line"  -ForegroundColor Cyan
    Write-Host '   Windows Developer Drive Cleaner — by win-cleaner' -ForegroundColor Gray
    Write-Host "  $line"  -ForegroundColor Cyan
    Write-Host ''
}

# ── Path prompt ────────────────────────────────────────────────────────────────
function Get-SearchPaths {
    param(
        [string[]]$Provided,
        [string[]]$DefaultPaths = @('D:\')
    )

    if ($Provided.Count -gt 0) {
        $valid = @()
        foreach ($p in $Provided) {
            if (Test-Path -LiteralPath $p) { $valid += $p }
            else { Write-Host "  WARNING: Path not found, skipping: $p" -ForegroundColor Yellow }
        }
        return $valid
    }

    $defaultLabel = $DefaultPaths -join ', '
    Write-Host 'Enter one or more root paths to scan.' -ForegroundColor Yellow
    Write-Host "Separate multiple paths with a comma. Press ENTER for default [$defaultLabel]:" -ForegroundColor Yellow
    $userInput = Read-Host '  Paths'

    if ([string]::IsNullOrWhiteSpace($userInput)) {
        $raw = $DefaultPaths
    }
    else {
        $raw = $userInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    }

    $valid = @()
    foreach ($p in $raw) {
        if (Test-Path -LiteralPath $p) { $valid += $p }
        else { Write-Host "  WARNING: Path not found, skipping: $p" -ForegroundColor Yellow }
    }

    if ($valid.Count -eq 0) {
        Write-Host 'No valid paths provided. Exiting.' -ForegroundColor Red
        exit 1
    }

    return $valid
}

# ── Step runner ────────────────────────────────────────────────────────────────
function Invoke-StepScript {
    param(
        [string]   $ScriptPath,
        [string[]] $Paths,
        [bool]     $DryRunMode,
        [hashtable]$ExtraParams = @{}
    )

    $pathArg = $Paths -join ','
    $args    = @{ SearchPaths = $Paths }
    if ($DryRunMode) { $args['DryRun'] = $true }
    foreach ($k in $ExtraParams.Keys) { $args[$k] = $ExtraParams[$k] }

    & $ScriptPath @args
}

# ── Menu ───────────────────────────────────────────────────────────────────────
function Show-Menu {
    param([string[]]$Paths, [bool]$DryRunMode)

    $dryTag = if ($DryRunMode) { ' [DRY RUN]' } else { '' }

    while ($true) {
        Write-Host ''
        Write-Host '─────────────────────────────────────────────────────────────' -ForegroundColor DarkCyan
        Write-Host "  WinCleaner Menu$dryTag" -ForegroundColor Cyan
        Write-Host "  Paths: $($Paths -join ', ')" -ForegroundColor Gray
        Write-Host '─────────────────────────────────────────────────────────────' -ForegroundColor DarkCyan
        Write-Host '  [1]  Step 1 — Dev Artifact Cleanup  (node_modules, dist, ...)' -ForegroundColor White
        Write-Host '  [2]  Step 2 — Duplicate File Cleanup' -ForegroundColor White
        Write-Host '  [3]  Step 3 — Empty Files & Folders Cleanup' -ForegroundColor White
        Write-Host '  [A]  Run ALL steps in sequence  (1 → 2 → 3)' -ForegroundColor Yellow
        Write-Host '  [P]  Change search paths' -ForegroundColor Gray
        Write-Host '  [Q]  Quit' -ForegroundColor Gray
        Write-Host '─────────────────────────────────────────────────────────────' -ForegroundColor DarkCyan

        $choice = (Read-Host '  Choose an option').Trim().ToUpper()

        switch ($choice) {
            '1' {
                Invoke-StepScript -ScriptPath $Step1Script -Paths $Paths -DryRunMode $DryRunMode
            }
            '2' {
                Invoke-StepScript -ScriptPath $Step2Script -Paths $Paths -DryRunMode $DryRunMode
            }
            '3' {
                Invoke-StepScript -ScriptPath $Step3Script -Paths $Paths -DryRunMode $DryRunMode
            }
            'A' {
                Write-Host ''
                Write-Host '══════════════════════════════════════════════════════════════' -ForegroundColor Yellow
                Write-Host '  Running ALL 3 Steps — Full Cleanup Sequence' -ForegroundColor Yellow
                Write-Host '══════════════════════════════════════════════════════════════' -ForegroundColor Yellow

                Write-Host ''
                Write-Host '  ┌─ Step 1/3 ─────────────────────────────────────────────┐' -ForegroundColor Cyan
                Invoke-StepScript -ScriptPath $Step1Script -Paths $Paths -DryRunMode $DryRunMode

                Write-Host ''
                Write-Host '  ┌─ Step 2/3 ─────────────────────────────────────────────┐' -ForegroundColor Cyan
                Invoke-StepScript -ScriptPath $Step2Script -Paths $Paths -DryRunMode $DryRunMode

                Write-Host ''
                Write-Host '  ┌─ Step 3/3 ─────────────────────────────────────────────┐' -ForegroundColor Cyan
                Invoke-StepScript -ScriptPath $Step3Script -Paths $Paths -DryRunMode $DryRunMode

                Write-Host ''
                Write-Host '══════════════════════════════════════════════════════════════' -ForegroundColor Green
                Write-Host '  All 3 steps completed. Check Reports\ folder for logs.' -ForegroundColor Green
                Write-Host '══════════════════════════════════════════════════════════════' -ForegroundColor Green
            }
            'P' {
                $Paths = Get-SearchPaths -Provided @()
                Write-Host "Paths updated to: $($Paths -join ', ')" -ForegroundColor Green
            }
            'Q' {
                Write-Host ''
                Write-Host 'WinCleaner exited. Reports saved in: ' -NoNewline -ForegroundColor Gray
                Write-Host (Join-Path $ScriptDir 'Reports') -ForegroundColor Cyan
                Write-Host ''
                return
            }
            default {
                Write-Host "Invalid option '$choice'. Please choose 1, 2, 3, A, P, or Q." -ForegroundColor Red
            }
        }
    }
}

# ── Entry point ────────────────────────────────────────────────────────────────
Show-Banner

$paths = Get-SearchPaths -Provided $SearchPaths -DefaultPaths @($Config.defaultScanPaths)

if ($DryRun) {
    Write-Host ''
    Write-Host '  *** DRY RUN MODE — Reports will be generated but nothing will be deleted. ***' -ForegroundColor Yellow
    Write-Host ''
}

if ($Step -ne '') {
    # Run a specific step directly (no menu)
    switch ($Step) {
        '1' { Invoke-StepScript -ScriptPath $Step1Script -Paths $paths -DryRunMode $DryRun.IsPresent }
        '2' { Invoke-StepScript -ScriptPath $Step2Script -Paths $paths -DryRunMode $DryRun.IsPresent }
        '3' { Invoke-StepScript -ScriptPath $Step3Script -Paths $paths -DryRunMode $DryRun.IsPresent }
    }
}
else {
    Show-Menu -Paths $paths -DryRunMode $DryRun.IsPresent
}
