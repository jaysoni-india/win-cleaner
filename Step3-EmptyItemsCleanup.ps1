#Requires -Version 5.1
<#
.SYNOPSIS
    Step 3 - Empty File and Empty Folder Cleanup

.DESCRIPTION
    Scans user-specified paths for:
      • Zero-byte files (0 KB files that are likely stale/abandoned)
      • Empty directories (no files anywhere in their subtree)

    Generates an HTML report, asks for confirmation, then deletes the
    empty items. Directories are deleted bottom-up so nested empty
    structures collapse cleanly.

.PARAMETER SearchPaths
    One or more root paths to scan. If omitted, the user is prompted.
    Default when prompted with no input: D:\

.PARAMETER SkipEmptyFiles
    If set, zero-byte files are not reported or deleted.

.PARAMETER SkipEmptyDirs
    If set, empty directories are not reported or deleted.

.PARAMETER DryRun
    If set, only generates the report without deleting anything.

.EXAMPLE
    .\Step3-EmptyItemsCleanup.ps1
    .\Step3-EmptyItemsCleanup.ps1 -SearchPaths "D:\Projects","E:\Work"
    .\Step3-EmptyItemsCleanup.ps1 -SearchPaths "D:\" -SkipEmptyFiles -DryRun
#>
[CmdletBinding()]
param(
    [string[]]$SearchPaths    = @(),
    [switch]  $SkipEmptyFiles,
    [switch]  $SkipEmptyDirs,
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

function Encode-Html {
    param([string]$s)
    $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

function Test-DirectoryEmpty {
    param([string]$Path)
    # A directory is empty if it contains no files anywhere in its subtree
    try {
        $anyFile = Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue |
                   Select-Object -First 1
        return ($null -eq $anyFile)
    }
    catch { return $false }
}

function New-Step3HtmlReport {
    param(
        [string]$SearchedPaths,
        [array] $EmptyFiles,
        [array] $EmptyDirs,
        [string]$OutputPath
    )

    $fileRows = foreach ($f in $EmptyFiles) {
        "<tr>
          <td>&#128196; Empty File</td>
          <td class='path'>$(Encode-Html $f.FullName)</td>
          <td class='size'>0 B</td>
          <td>$($f.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))</td>
        </tr>"
    }

    $dirRows = foreach ($d in $EmptyDirs) {
        "<tr>
          <td>&#128193; Empty Folder</td>
          <td class='path'>$(Encode-Html $d.FullName)</td>
          <td class='size'>—</td>
          <td>$($d.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))</td>
        </tr>"
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>WinCleaner - Step 3 Report</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;padding:24px;color:#333}
    .wrap{max-width:1400px;margin:0 auto}
    h1{color:#27ae60;font-size:24px;margin-bottom:4px}
    .sub{color:#777;font-size:13px;margin-bottom:22px}
    .warn{background:#d4edda;border:1px solid #c3e6cb;border-radius:6px;padding:12px 16px;margin-bottom:20px;font-size:13px}
    .card{background:#fff;border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.1);padding:20px;margin-bottom:20px}
    .stats{display:flex;gap:40px;flex-wrap:wrap;margin-top:10px}
    .stat .num{font-size:32px;font-weight:700;color:#27ae60}
    .stat .lbl{font-size:12px;color:#888;margin-top:2px}
    table{width:100%;border-collapse:collapse;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.1)}
    th{background:#2c3e50;color:#fff;padding:11px 14px;text-align:left;font-size:13px}
    td{padding:9px 14px;border-bottom:1px solid #eee;font-size:12px;vertical-align:middle}
    tr:last-child td{border-bottom:none}
    tr:hover td{background:#f0faf2}
    .path{font-family:Consolas,monospace;font-size:11px;word-break:break-all}
    .size{text-align:right;font-family:Consolas,monospace;white-space:nowrap}
    .section{background:#ecf9ef;font-weight:600;color:#1a5c2e;padding:9px 14px;border-bottom:2px solid #a9dfbf;font-size:12px}
    .footer{text-align:center;margin-top:24px;font-size:11px;color:#bbb}
  </style>
</head>
<body>
<div class="wrap">
  <h1>&#9855; WinCleaner &mdash; Step 3: Empty Files &amp; Folders</h1>
  <p class="sub">Search paths: $(Encode-Html $SearchedPaths) &bull; Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
  <div class="warn">&#8505;&#65039; All items listed below are empty and safe to remove. Confirm deletion in the terminal after reviewing this report.</div>
  <div class="card">
    <strong>Scan Summary</strong>
    <div class="stats">
      <div class="stat"><div class="num">$($EmptyFiles.Count)</div><div class="lbl">Zero-byte Files</div></div>
      <div class="stat"><div class="num">$($EmptyDirs.Count)</div><div class="lbl">Empty Folders</div></div>
      <div class="stat"><div class="num">$($EmptyFiles.Count + $EmptyDirs.Count)</div><div class="lbl">Total Items</div></div>
    </div>
  </div>
  <table>
    <thead>
      <tr><th>Type</th><th>Full Path</th><th>Size</th><th>Last Modified</th></tr>
    </thead>
    <tbody>
      $(if ($EmptyFiles.Count -gt 0) { '<tr class="section"><td colspan="4">&#128196; Zero-Byte Files (' + $EmptyFiles.Count + ')</td></tr>' + ($fileRows -join "`n") })
      $(if ($EmptyDirs.Count  -gt 0) { '<tr class="section"><td colspan="4">&#128193; Empty Folders ('  + $EmptyDirs.Count  + ')</td></tr>' + ($dirRows  -join "`n") })
    </tbody>
  </table>
  <div class="footer">WinCleaner &bull; Step 3 &bull; Empty Items Cleanup</div>
</div>
</body>
</html>
"@
    $html | Set-Content -Path $OutputPath -Encoding UTF8
}

# ── Main logic ─────────────────────────────────────────────────────────────────

function Invoke-Step3 {
    param(
        [string[]]$Paths,
        [bool]    $NoEmptyFiles,
        [bool]    $NoEmptyDirs,
        [bool]    $DryRunMode
    )

    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host '  STEP 3 — Empty Files & Empty Folder Cleanup' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host ''

    # ── Resolve search paths ──────────────────────────────────────────────────
    if ($Paths.Count -eq 0) {
        Write-Host 'Enter search paths (comma-separated). Press ENTER for default [D:\]:' -ForegroundColor Yellow
        $userInput = Read-Host '  Paths'
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            $Paths = @('D:\')
        }
        else {
            $Paths = $userInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
        }
    }

    $validPaths = @()
    foreach ($p in $Paths) {
        if (Test-Path -LiteralPath $p) { $validPaths += $p }
        else { Write-Host "  WARNING: Path not found, skipping: $p" -ForegroundColor Yellow }
    }
    if ($validPaths.Count -eq 0) {
        Write-Host 'No valid paths to search. Exiting.' -ForegroundColor Red
        return
    }

    Write-Host "Searching in : $($validPaths -join ', ')" -ForegroundColor Cyan
    Write-Host 'Please wait...'
    Write-Host ''

    # ── Scan for zero-byte files ──────────────────────────────────────────────
    $emptyFiles = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

    if (-not $NoEmptyFiles) {
        Write-Progress -Activity 'Step 3 - Scanning' -Status 'Finding zero-byte files...' -PercentComplete 25

        foreach ($rootPath in $validPaths) {
            try {
                Get-ChildItem -LiteralPath $rootPath -Recurse -Force -File -ErrorAction SilentlyContinue |
                    Where-Object { $_.Length -eq 0 } |
                    ForEach-Object { $emptyFiles.Add($_) }
            }
            catch {
                Write-Host "  Warning: Error scanning $rootPath — $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }

    # ── Scan for empty directories ────────────────────────────────────────────
    $emptyDirs = [System.Collections.Generic.List[System.IO.DirectoryInfo]]::new()

    if (-not $NoEmptyDirs) {
        Write-Progress -Activity 'Step 3 - Scanning' -Status 'Finding empty directories...' -PercentComplete 60

        foreach ($rootPath in $validPaths) {
            try {
                # Get all directories, sort by depth descending (deepest first)
                # so we process leaves before parents
                $allDirs = Get-ChildItem -LiteralPath $rootPath -Recurse -Force -Directory -ErrorAction SilentlyContinue |
                    Sort-Object { $_.FullName.Split('\').Count } -Descending

                foreach ($dir in $allDirs) {
                    if (Test-DirectoryEmpty -Path $dir.FullName) {
                        $emptyDirs.Add($dir)
                    }
                }
            }
            catch {
                Write-Host "  Warning: Error scanning $rootPath — $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        # Remove duplicates that are sub-paths of other empty dirs already in list
        # (parent will be deleted when child is deleted; avoid double-reporting)
        $dedupedDirs = [System.Collections.Generic.List[System.IO.DirectoryInfo]]::new()
        foreach ($d in $emptyDirs) {
            $isChild = $false
            foreach ($other in $emptyDirs) {
                if ($d.FullName -ne $other.FullName -and
                    $d.FullName.StartsWith($other.FullName + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
                    $isChild = $true
                    break
                }
            }
            if (-not $isChild) { $dedupedDirs.Add($d) }
        }
        $emptyDirs = $dedupedDirs
    }

    Write-Progress -Activity 'Step 3 - Scanning' -Completed

    $totalItems = $emptyFiles.Count + $emptyDirs.Count

    if ($totalItems -eq 0) {
        Write-Host 'No empty files or folders found. Your workspace is spotless!' -ForegroundColor Green
        return
    }

    Write-Host "Zero-byte files : $($emptyFiles.Count)" -ForegroundColor Yellow
    Write-Host "Empty folders   : $($emptyDirs.Count)"  -ForegroundColor Yellow
    Write-Host "Total items     : $totalItems"           -ForegroundColor Yellow
    Write-Host ''

    # ── Generate HTML report ──────────────────────────────────────────────────
    $reportDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path $reportDir)) { New-Item -ItemType Directory -Path $reportDir | Out-Null }

    $ts         = Get-Date -Format 'yyyyMMdd-HHmmss'
    $reportPath = Join-Path $reportDir "Step3-EmptyItems-$ts.html"

    New-Step3HtmlReport `
        -SearchedPaths ($validPaths -join ', ') `
        -EmptyFiles    $emptyFiles `
        -EmptyDirs     $emptyDirs `
        -OutputPath    $reportPath

    Write-Host "Report saved : $reportPath" -ForegroundColor Green

    $openAns = Read-Host 'Open report in browser now? (Y/N) [Y]'
    if ([string]::IsNullOrWhiteSpace($openAns) -or $openAns -imatch '^y') {
        Start-Process $reportPath
        Write-Host 'Waiting 5 seconds for you to review the report...' -ForegroundColor DarkCyan
        Start-Sleep -Seconds 5
    }

    Write-Host ''

    if ($DryRunMode) {
        Write-Host 'DRY RUN mode — no files were deleted.' -ForegroundColor Cyan
        return
    }

    # ── Confirm deletion ──────────────────────────────────────────────────────
    Write-Host 'Type  DELETE  to confirm removal of all empty items, or press ENTER to cancel:' -ForegroundColor Red
    $confirm = Read-Host '  Confirm'

    if ($confirm -cne 'DELETE') {
        Write-Host 'Deletion cancelled. No files were removed.' -ForegroundColor Cyan
        return
    }

    # ── Perform deletion ──────────────────────────────────────────────────────
    $logPath = Join-Path $reportDir "Step3-DeletionLog-$ts.txt"
    "WinCleaner Step 3 — Deletion Log"       | Set-Content  $logPath -Encoding UTF8
    "Date   : $(Get-Date)"                   | Add-Content  $logPath
    "Paths  : $($validPaths -join ', ')"     | Add-Content  $logPath
    ('─' * 60)                               | Add-Content  $logPath

    $deleted = 0
    $errored = 0

    # Delete zero-byte files first
    foreach ($f in $emptyFiles) {
        if (-not (Test-Path -LiteralPath $f.FullName)) { continue }
        try {
            Remove-Item -LiteralPath $f.FullName -Force
            "[DELETED FILE] $($f.FullName)" | Add-Content $logPath
            $deleted++
            Write-Host "  Deleted file: $($f.FullName)" -ForegroundColor DarkGray
        }
        catch {
            "[ERROR FILE]   $($f.FullName) — $($_.Exception.Message)" | Add-Content $logPath
            $errored++
            Write-Host "  ERROR  : $($f.FullName)" -ForegroundColor Red
        }
    }

    # Delete empty directories (already deduplicated; Remove-Item -Recurse handles any nested empties)
    foreach ($d in $emptyDirs) {
        if (-not (Test-Path -LiteralPath $d.FullName)) { continue }
        # Re-confirm still empty (zero-byte file deletions above may have affected state)
        if (-not (Test-DirectoryEmpty -Path $d.FullName)) { continue }
        try {
            Remove-Item -LiteralPath $d.FullName -Recurse -Force
            "[DELETED DIR]  $($d.FullName)" | Add-Content $logPath
            $deleted++
            Write-Host "  Deleted dir : $($d.FullName)" -ForegroundColor DarkGray
        }
        catch {
            "[ERROR DIR]    $($d.FullName) — $($_.Exception.Message)" | Add-Content $logPath
            $errored++
            Write-Host "  ERROR  : $($d.FullName)" -ForegroundColor Red
        }
    }

    Write-Host ''
    Write-Host '── Step 3 Complete ───────────────────────────────────────────' -ForegroundColor Green
    Write-Host "   Deleted : $deleted items"                                    -ForegroundColor Green
    if ($errored -gt 0) {
        Write-Host "   Errors  : $errored items (check log)"                    -ForegroundColor Yellow
    }
    Write-Host "   Log     : $logPath"                                          -ForegroundColor Gray
    Write-Host ''
}

# ── Entry point ────────────────────────────────────────────────────────────────
if ($MyInvocation.InvocationName -ne '.') {
    Invoke-Step3 -Paths         $SearchPaths `
                 -NoEmptyFiles  $SkipEmptyFiles.IsPresent `
                 -NoEmptyDirs   $SkipEmptyDirs.IsPresent `
                 -DryRunMode    $DryRun.IsPresent
}
