#Requires -Version 5.1
<#
.SYNOPSIS
    Step 1 - Dev Artifact Folder/File Cleanup

.DESCRIPTION
    Scans user-specified paths for developer artifact folders and files defined
    in cleanup-list.txt (node_modules, .gradle, __pycache__, etc.).
    Generates an HTML report showing what will be deleted (with sizes), asks for
    confirmation, then performs deletion and saves a deletion log.

.PARAMETER SearchPaths
    One or more root paths to scan. If omitted, the user is prompted.
    Default when prompted with no input: D:\

.PARAMETER CleanupListFile
    Path to cleanup-list.txt. Defaults to cleanup-list.txt in the same folder
    as this script.

.PARAMETER DryRun
    If set, only generates the report without deleting anything.

.EXAMPLE
    .\Step1-DevFolderCleanup.ps1
    .\Step1-DevFolderCleanup.ps1 -SearchPaths "D:\Projects","E:\Work"
    .\Step1-DevFolderCleanup.ps1 -SearchPaths "D:\" -DryRun
#>
[CmdletBinding()]
param(
    [string[]]$SearchPaths   = @(),
    [string]  $CleanupListFile = '',
    [switch]  $DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# ── Helper functions ───────────────────────────────────────────────────────────

function Format-Size {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N2} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N2} KB' -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Get-FolderSize {
    param([string]$Path)
    try {
        $sum = (Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
                Where-Object { -not $_.PSIsContainer } |
                Measure-Object -Property Length -Sum).Sum
        return [long]$(if ($null -eq $sum) { 0 } else { $sum })
    }
    catch { return 0L }
}

function Get-WinCleanerConfig {
    param([string]$ScriptRoot)
    $configFile = Join-Path $ScriptRoot 'config.json'
    $defaults = [PSCustomObject]@{
        defaultScanPaths = @('D:\')
        ignoreNames      = @('.git', '.env')
        ignorePaths      = @()
    }
    if (-not (Test-Path -LiteralPath $configFile)) { return $defaults }
    try {
        $cfg = Get-Content -LiteralPath $configFile -Raw -Encoding UTF8 | ConvertFrom-Json
        # Ensure all keys exist, fall back to defaults for missing ones
        if ($null -eq $cfg.defaultScanPaths -or $cfg.defaultScanPaths.Count -eq 0) {
            $cfg | Add-Member -Force -NotePropertyName defaultScanPaths -NotePropertyValue $defaults.defaultScanPaths
        }
        if ($null -eq $cfg.ignoreNames)  { $cfg | Add-Member -Force -NotePropertyName ignoreNames  -NotePropertyValue $defaults.ignoreNames }
        if ($null -eq $cfg.ignorePaths)  { $cfg | Add-Member -Force -NotePropertyName ignorePaths  -NotePropertyValue $defaults.ignorePaths }
        return $cfg
    }
    catch {
        Write-Host "  WARNING: Could not parse config.json — using defaults. ($($_.Exception.Message))" -ForegroundColor Yellow
        return $defaults
    }
}

function Get-CleanupPatterns {
    param([string]$File)
    $list = @()
    foreach ($line in (Get-Content -Path $File -Encoding UTF8)) {
        $t = $line.Trim()
        if ($t -and -not $t.StartsWith('#')) { $list += $t }
    }
    return $list
}

function Resolve-MatchedPattern {
    param([string]$Name, [string[]]$Patterns)
    foreach ($p in $Patterns) {
        if ($Name -like $p) { return $p }
    }
    return $null
}

function Encode-Html {
    param([string]$s)
    $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

function New-Step1HtmlReport {
    param(
        [string]$SearchedPaths,
        [array] $Items,
        [long]  $TotalBytes,
        [string]$OutputPath
    )

    $rowsHtml = foreach ($item in $Items) {
        $icon = if ($item.Type -eq 'Directory') { '&#128193;' } else { '&#128196;' }
        "<tr>
          <td>$icon $(Encode-Html $item.Type)</td>
          <td><code>$(Encode-Html $item.MatchedPattern)</code></td>
          <td class='path'>$(Encode-Html $item.FullName)</td>
          <td class='size'>$($item.SizeFormatted)</td>
          <td>$($item.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))</td>
        </tr>"
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>WinCleaner - Step 1 Report</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;padding:24px;color:#333}
    .wrap{max-width:1400px;margin:0 auto}
    h1{color:#c0392b;font-size:24px;margin-bottom:4px}
    .sub{color:#777;font-size:13px;margin-bottom:22px}
    .warn{background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:12px 16px;margin-bottom:20px;font-size:13px}
    .card{background:#fff;border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.1);padding:20px;margin-bottom:20px}
    .stats{display:flex;gap:40px;flex-wrap:wrap;margin-top:10px}
    .stat .num{font-size:32px;font-weight:700;color:#c0392b}
    .stat .lbl{font-size:12px;color:#888;margin-top:2px}
    table{width:100%;border-collapse:collapse;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.1)}
    th{background:#2c3e50;color:#fff;padding:11px 14px;text-align:left;font-size:13px}
    td{padding:9px 14px;border-bottom:1px solid #eee;font-size:12px;vertical-align:middle}
    tr:last-child td{border-bottom:none}
    tr:hover td{background:#fef9f9}
    .path{font-family:Consolas,monospace;font-size:11px;word-break:break-all}
    .size{text-align:right;font-family:Consolas,monospace;white-space:nowrap}
    code{background:#f0f0f0;padding:2px 6px;border-radius:3px;font-size:11px;font-family:Consolas,monospace}
    .footer{text-align:center;margin-top:24px;font-size:11px;color:#bbb}
  </style>
</head>
<body>
<div class="wrap">
  <h1>&#128465; WinCleaner &mdash; Step 1: Dev Artifact Cleanup</h1>
  <p class="sub">Search paths: $(Encode-Html $SearchedPaths) &bull; Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
  <div class="warn">&#9888;&#65039; <strong>Review carefully before confirming deletion.</strong> All items below will be permanently removed.</div>
  <div class="card">
    <strong>Scan Summary</strong>
    <div class="stats">
      <div class="stat"><div class="num">$($Items.Count)</div><div class="lbl">Items Found</div></div>
      <div class="stat"><div class="num">$(Format-Size $TotalBytes)</div><div class="lbl">Disk Space to Reclaim</div></div>
    </div>
  </div>
  <table>
    <thead>
      <tr><th>Type</th><th>Pattern Matched</th><th>Full Path</th><th>Size</th><th>Last Modified</th></tr>
    </thead>
    <tbody>
      $($rowsHtml -join "`n")
    </tbody>
  </table>
  <div class="footer">WinCleaner &bull; Step 1 &bull; Dev Artifact Cleanup</div>
</div>
</body>
</html>
"@
    $html | Set-Content -Path $OutputPath -Encoding UTF8
}

# ── Main logic ─────────────────────────────────────────────────────────────────

function Invoke-Step1 {
    param(
        [string[]]$Paths,
        [string]  $ListFile,
        [bool]    $DryRunMode
    )

    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host '  STEP 1 — Dev Artifact Folder / File Cleanup' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host ''

    # ── Load config.json ──────────────────────────────────────────────────────
    $Config = Get-WinCleanerConfig -ScriptRoot $PSScriptRoot

    # ── Resolve cleanup list ──────────────────────────────────────────────────
    if ([string]::IsNullOrWhiteSpace($ListFile)) {
        $ListFile = Join-Path $PSScriptRoot 'cleanup-list.txt'
    }
    if (-not (Test-Path -LiteralPath $ListFile)) {
        Write-Host "ERROR: cleanup-list.txt not found at: $ListFile" -ForegroundColor Red
        return
    }

    # ── Resolve search paths ──────────────────────────────────────────────────
    if ($Paths.Count -eq 0) {
        $defaultPaths = @($Config.defaultScanPaths)
        $defaultLabel = $defaultPaths -join ', '
        Write-Host "Enter search paths (comma-separated). Press ENTER for default [$defaultLabel]:" -ForegroundColor Yellow
        $userInput = Read-Host '  Paths'
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            $Paths = $defaultPaths
        }
        else {
            $Paths = $userInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
        }
    }

    $validPaths = @()
    foreach ($p in $Paths) {
        if (Test-Path -LiteralPath $p) {
            $validPaths += $p
        }
        else {
            Write-Host "  WARNING: Path not found, skipping: $p" -ForegroundColor Yellow
        }
    }
    if ($validPaths.Count -eq 0) {
        Write-Host 'No valid paths to search. Exiting.' -ForegroundColor Red
        return
    }

    # ── Ignore list from config.json — never deleted or reported ─────────────
    # ignoreNames: exact folder/file names to skip (e.g. .git, .env)
    # ignorePaths: full path prefixes/substrings to skip (e.g. C:\sensitive)
    $IgnoreNames = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($Config.ignoreNames),
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $IgnorePaths = @($Config.ignorePaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    # ── Load patterns ─────────────────────────────────────────────────────────
    $patterns = Get-CleanupPatterns -File $ListFile
    Write-Host "Loaded $($patterns.Count) cleanup patterns." -ForegroundColor Green
    Write-Host "Ignored names : $($IgnoreNames -join ', ')" -ForegroundColor DarkGray
    if ($IgnorePaths.Count -gt 0) {
        Write-Host "Ignored paths : $($IgnorePaths -join ', ')" -ForegroundColor DarkGray
    }
    Write-Host "Searching in: $($validPaths -join ', ')" -ForegroundColor Cyan
    Write-Host 'Please wait...'
    Write-Host ''

    # ── Scan ──────────────────────────────────────────────────────────────────
    $found    = [System.Collections.Generic.List[PSObject]]::new()
    $skipSet  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $scanCount = 0

    foreach ($rootPath in $validPaths) {
        Write-Progress -Activity 'Step 1 - Scanning for dev artifacts' -Status "Scanning: $rootPath" -PercentComplete -1

        try {
            $allItems = Get-ChildItem -LiteralPath $rootPath -Recurse -Force -ErrorAction SilentlyContinue

            foreach ($item in $allItems) {
                $scanCount++
                if ($scanCount % 500 -eq 0) {
                    Write-Progress -Activity 'Step 1 - Scanning for dev artifacts' `
                        -Status "Scanned $scanCount items | Found $($found.Count) matches" `
                        -PercentComplete -1
                }

                # Skip items inside already-matched parent folders
                $isNested = $false
                foreach ($skip in $skipSet) {
                    if ($item.FullName.StartsWith($skip + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
                        $isNested = $true
                        break
                    }
                }
                if ($isNested) { continue }

                # Skip names listed in config.json ignoreNames (.git, .env, etc.)
                if ($IgnoreNames.Contains($item.Name)) { continue }
                # Skip anything whose full path is inside an ignored-name directory
                # (e.g. every file/folder inside any .git directory)
                $insideIgnored = $false
                foreach ($ignoreName in $IgnoreNames) {
                    if ($item.FullName -like "*\$ignoreName\*") { $insideIgnored = $true; break }
                }
                if ($insideIgnored) { continue }
                # Skip paths matching any entry in config.json ignorePaths
                foreach ($ignorePath in $IgnorePaths) {
                    if ($item.FullName.StartsWith($ignorePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $insideIgnored = $true; break
                    }
                }
                if ($insideIgnored) { continue }

                $matched = Resolve-MatchedPattern -Name $item.Name -Patterns $patterns
                if ($null -ne $matched) {
                    $size = if ($item.PSIsContainer) { Get-FolderSize $item.FullName } else { $item.Length }

                    $found.Add([PSCustomObject]@{
                        Type           = if ($item.PSIsContainer) { 'Directory' } else { 'File' }
                        Name           = $item.Name
                        FullName       = $item.FullName
                        MatchedPattern = $matched
                        SizeBytes      = $size
                        SizeFormatted  = Format-Size $size
                        LastWriteTime  = $item.LastWriteTime
                    })

                    if ($item.PSIsContainer) { [void]$skipSet.Add($item.FullName) }
                }
            }
        }
        catch {
            Write-Host "  Warning: Error scanning $rootPath — $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    Write-Progress -Activity 'Step 1 - Scanning for dev artifacts' -Completed

    if ($found.Count -eq 0) {
        Write-Host 'No matching items found. Your workspace is already clean!' -ForegroundColor Green
        return
    }

    $totalBytes = ($found | Measure-Object -Property SizeBytes -Sum).Sum

    Write-Host "Found : $($found.Count) items matching cleanup patterns" -ForegroundColor Yellow
    Write-Host "Size  : $(Format-Size $totalBytes) will be reclaimed" -ForegroundColor Yellow
    Write-Host ''

    # ── Generate HTML report ──────────────────────────────────────────────────
    $reportDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path $reportDir)) { New-Item -ItemType Directory -Path $reportDir | Out-Null }

    $ts         = Get-Date -Format 'yyyyMMdd-HHmmss'
    $reportPath = Join-Path $reportDir "Step1-DevArtifacts-$ts.html"

    New-Step1HtmlReport `
        -SearchedPaths ($validPaths -join ', ') `
        -Items         $found `
        -TotalBytes    $totalBytes `
        -OutputPath    $reportPath

    Write-Host "Report saved : $reportPath" -ForegroundColor Green

    $openAns = Read-Host 'Open report in browser now? (Y/N) [Y]'
    if ([string]::IsNullOrWhiteSpace($openAns) -or $openAns -imatch '^y') {
        Start-Process $reportPath
        Write-Host 'Waiting 5 seconds for you to review the report...' -ForegroundColor DarkCyan
        Start-Sleep -Seconds 5
    }

    Write-Host ''

    # ── Dry-run guard ─────────────────────────────────────────────────────────
    if ($DryRunMode) {
        Write-Host 'DRY RUN mode — no files were deleted.' -ForegroundColor Cyan
        return
    }

    # ── Confirm deletion ──────────────────────────────────────────────────────
    Write-Host 'After reviewing the report, confirm permanent deletion.' -ForegroundColor Yellow
    Write-Host "Type  DELETE  to proceed, or press ENTER to cancel:" -ForegroundColor Red
    $confirm = Read-Host '  Confirm'

    if ($confirm -cne 'DELETE') {
        Write-Host 'Deletion cancelled. No files were removed.' -ForegroundColor Cyan
        return
    }

    # ── Perform deletion ──────────────────────────────────────────────────────
    $logPath = Join-Path $reportDir "Step1-DeletionLog-$ts.txt"
    "WinCleaner Step 1 — Deletion Log"       | Set-Content  $logPath -Encoding UTF8
    "Date   : $(Get-Date)"                   | Add-Content  $logPath
    "Paths  : $($validPaths -join ', ')"     | Add-Content  $logPath
    ('─' * 60)                               | Add-Content  $logPath

    $deleted = 0
    $errored = 0

    foreach ($item in $found) {
        if (-not (Test-Path -LiteralPath $item.FullName)) { continue } # already gone (nested)

        try {
            if ($item.Type -eq 'Directory') {
                Remove-Item -LiteralPath $item.FullName -Recurse -Force
            }
            else {
                Remove-Item -LiteralPath $item.FullName -Force
            }
            "[DELETED] $($item.FullName) ($($item.SizeFormatted))" | Add-Content $logPath
            $deleted++
            Write-Host "  Deleted: $($item.FullName)" -ForegroundColor DarkGray
        }
        catch {
            "[ERROR]   $($item.FullName) — $($_.Exception.Message)" | Add-Content $logPath
            $errored++
            Write-Host "  ERROR  : $($item.FullName)" -ForegroundColor Red
        }
    }

    Write-Host ''
    Write-Host '── Step 1 Complete ───────────────────────────────────────────' -ForegroundColor Green
    Write-Host "   Deleted : $deleted items"                                    -ForegroundColor Green
    if ($errored -gt 0) {
        Write-Host "   Errors  : $errored items (check log for details)"        -ForegroundColor Yellow
    }
    Write-Host "   Log     : $logPath"                                          -ForegroundColor Gray
    Write-Host ''
}

# ── Entry point ────────────────────────────────────────────────────────────────
# Only run automatically when executed directly (not dot-sourced by WinCleaner.ps1)
if ($MyInvocation.InvocationName -ne '.') {
    Invoke-Step1 -Paths $SearchPaths -ListFile $CleanupListFile -DryRunMode $DryRun.IsPresent
}
