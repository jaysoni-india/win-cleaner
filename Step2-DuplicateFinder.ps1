#Requires -Version 5.1
<#
.SYNOPSIS
    Step 2 - Duplicate File Finder & Cleaner

.DESCRIPTION
    Scans user-specified paths for duplicate files (identical content detected via
    SHA-256 hash). Groups duplicates, generates an HTML report showing every group
    with file sizes and paths, then asks the user how to resolve each group before
    deleting.

    Resolution options per group:
      K = Keep newest (delete older copies)   [default]
      O = Keep oldest (delete newer copies)
      S = Skip this group (do not delete)

.PARAMETER SearchPaths
    One or more root paths to scan. If omitted, the user is prompted.
    Default when prompted with no input: D:\

.PARAMETER MinSizeKB
    Minimum file size in KB to consider. Files smaller than this are ignored.
    Default: 1 (1 KB). Set to 0 to include all files.

.PARAMETER DryRun
    If set, only generates the report without deleting anything.

.EXAMPLE
    .\Step2-DuplicateFinder.ps1
    .\Step2-DuplicateFinder.ps1 -SearchPaths "D:\Projects","E:\Work" -MinSizeKB 10
    .\Step2-DuplicateFinder.ps1 -SearchPaths "D:\" -DryRun
#>
[CmdletBinding()]
param(
    [string[]]$SearchPaths = @(),
    [int]     $MinSizeKB   = 1,
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

function Get-FileHashSafe {
    param([string]$Path)
    try {
        return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
    }
    catch { return $null }
}

function Get-HashCacheTable {
    param([string]$CacheFile)
    if (-not (Test-Path -LiteralPath $CacheFile)) { return @{} }
    try {
        $json  = Get-Content -LiteralPath $CacheFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $table = @{}
        foreach ($prop in $json.PSObject.Properties) {
            $table[$prop.Name] = $prop.Value
        }
        return $table
    }
    catch { return @{} }
}

function Save-HashCacheTable {
    param([hashtable]$Cache, [string]$CacheFile)
    try {
        $obj = [ordered]@{}
        foreach ($key in ($Cache.Keys | Sort-Object)) { $obj[$key] = $Cache[$key] }
        $obj | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $CacheFile -Encoding UTF8
    }
    catch {
        Write-Host "  WARNING: Could not save hash cache: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function New-Step2HtmlReport {
    param(
        [string]$SearchedPaths,
        [array] $DuplicateGroups,
        [long]  $WasteBytes,
        [string]$OutputPath
    )

    $groupRows = foreach ($g in $DuplicateGroups) {
        $fileRows = foreach ($f in $g.Files) {
            $keep = if ($f.FullName -eq $g.KeepFile) { '<span style="color:#27ae60;font-weight:bold">&#10003; KEEP</span>' } else { '<span style="color:#c0392b">&#10005; DELETE</span>' }
            "<tr>
              <td>$keep</td>
              <td class='path'>$(Encode-Html $f.FullName)</td>
              <td class='size'>$(Format-Size $f.Length)</td>
              <td>$($f.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))</td>
            </tr>"
        }
        "<tr class='group-header'>
          <td colspan='4'>&#128203; Group $(Encode-Html $g.GroupId) &bull; Hash: <code>$($g.Hash.Substring(0,16))&hellip;</code> &bull; $($g.Files.Count) copies &bull; Wasted: $(Format-Size $g.WasteBytes)</td>
        </tr>
        $($fileRows -join "`n")"
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>WinCleaner - Step 2 Report</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;padding:24px;color:#333}
    .wrap{max-width:1400px;margin:0 auto}
    h1{color:#2980b9;font-size:24px;margin-bottom:4px}
    .sub{color:#777;font-size:13px;margin-bottom:22px}
    .warn{background:#d1ecf1;border:1px solid #bee5eb;border-radius:6px;padding:12px 16px;margin-bottom:20px;font-size:13px}
    .card{background:#fff;border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.1);padding:20px;margin-bottom:20px}
    .stats{display:flex;gap:40px;flex-wrap:wrap;margin-top:10px}
    .stat .num{font-size:32px;font-weight:700;color:#2980b9}
    .stat .lbl{font-size:12px;color:#888;margin-top:2px}
    table{width:100%;border-collapse:collapse;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.1)}
    th{background:#2c3e50;color:#fff;padding:11px 14px;text-align:left;font-size:13px}
    td{padding:8px 14px;border-bottom:1px solid #eee;font-size:12px;vertical-align:middle}
    .group-header td{background:#eaf2fb;font-size:12px;font-weight:600;color:#1a5276;padding:8px 14px;border-bottom:2px solid #aed6f1}
    .path{font-family:Consolas,monospace;font-size:11px;word-break:break-all}
    .size{text-align:right;font-family:Consolas,monospace;white-space:nowrap}
    code{background:#f0f0f0;padding:2px 6px;border-radius:3px;font-size:11px;font-family:Consolas,monospace}
    .footer{text-align:center;margin-top:24px;font-size:11px;color:#bbb}
  </style>
</head>
<body>
<div class="wrap">
  <h1>&#128203; WinCleaner &mdash; Step 2: Duplicate File Finder</h1>
  <p class="sub">Search paths: $(Encode-Html $SearchedPaths) &bull; Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
  <div class="warn">&#8505;&#65039; Files marked <strong>KEEP</strong> will be retained (newest by default). All other copies will be deleted after your confirmation in the terminal.</div>
  <div class="card">
    <strong>Scan Summary</strong>
    <div class="stats">
      <div class="stat"><div class="num">$($DuplicateGroups.Count)</div><div class="lbl">Duplicate Groups</div></div>
      <div class="stat"><div class="num">$(Format-Size $WasteBytes)</div><div class="lbl">Reclaimable Space</div></div>
    </div>
  </div>
  <table>
    <thead>
      <tr><th>Action</th><th>Full Path</th><th>Size</th><th>Last Modified</th></tr>
    </thead>
    <tbody>
      $($groupRows -join "`n")
    </tbody>
  </table>
  <div class="footer">WinCleaner &bull; Step 2 &bull; Duplicate File Finder</div>
</div>
</body>
</html>
"@
    $html | Set-Content -Path $OutputPath -Encoding UTF8
}

# ── Main logic ─────────────────────────────────────────────────────────────────

function Invoke-Step2 {
    param(
        [string[]]$Paths,
        [int]     $MinKB,
        [bool]    $DryRunMode
    )

    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host '  STEP 2 — Duplicate File Finder & Cleanup' -ForegroundColor Cyan
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

    $minBytes = [long]$MinKB * 1024L

    Write-Host "Searching in : $($validPaths -join ', ')" -ForegroundColor Cyan
    Write-Host "Min file size: $(Format-Size $minBytes)" -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'Phase 1/2 — Collecting files...' -ForegroundColor DarkCyan

    # ── Phase 1: collect files, group by size ─────────────────────────────────
    $allFiles = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

    foreach ($rootPath in $validPaths) {
        try {
            Get-ChildItem -LiteralPath $rootPath -Recurse -Force -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Length -ge $minBytes } |
                ForEach-Object { $allFiles.Add($_) }
        }
        catch {
            Write-Host "  Warning: Error scanning $rootPath — $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    Write-Host "Found $($allFiles.Count) files (>= $(Format-Size $minBytes))." -ForegroundColor Gray

    # Group by size — only groups with 2+ files are potential duplicates
    $bySizeGroups = $allFiles |
        Group-Object -Property Length |
        Where-Object { $_.Count -gt 1 }

    Write-Host "Size-based candidates: $($bySizeGroups.Count) groups ($( ($bySizeGroups | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum ) files)."
    Write-Host ''
    Write-Host 'Phase 2/2 — Computing SHA-256 hashes for candidate files...' -ForegroundColor DarkCyan

    # ── Phase 2: hash candidates, find true duplicates ────────────────────────
    # Load persistent hash cache — avoids rehashing unchanged files on repeat scans
    $cacheFile  = Join-Path $PSScriptRoot 'hash-cache.json'
    $hashCache  = Get-HashCacheTable -CacheFile $cacheFile
    $cacheHits  = 0
    $cacheMisses = 0

    $hashMap   = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[System.IO.FileInfo]]]::new()
    $hashCount = 0
    $totalCandidates = ($bySizeGroups | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum

    foreach ($sizeGroup in $bySizeGroups) {
        foreach ($fi in $sizeGroup.Group) {
            $hashCount++
            if ($hashCount % 50 -eq 0) {
                Write-Progress -Activity 'Step 2 - Hashing files' `
                    -Status "Hashed $hashCount / $totalCandidates  (cache hits: $cacheHits)" `
                    -PercentComplete ([int](($hashCount / [Math]::Max($totalCandidates, 1)) * 100))
            }

            # Cache lookup: valid when path, file size and last-write time all match
            $cacheKey = $fi.FullName
            $fileLwt  = $fi.LastWriteTime.ToString('o')
            $hash     = $null

            if ($hashCache.ContainsKey($cacheKey)) {
                $entry = $hashCache[$cacheKey]
                if ([long]$entry.Size -eq $fi.Length -and $entry.LastWriteTime -eq $fileLwt) {
                    $hash = $entry.Hash
                    $cacheHits++
                }
            }

            if ($null -eq $hash) {
                $hash = Get-FileHashSafe -Path $fi.FullName
                if ($null -eq $hash) { continue }
                # Store in cache for future runs
                $hashCache[$cacheKey] = [PSCustomObject]@{
                    Hash          = $hash
                    Size          = $fi.Length
                    LastWriteTime = $fileLwt
                }
                $cacheMisses++
            }

            if (-not $hashMap.ContainsKey($hash)) {
                $hashMap[$hash] = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
            }
            $hashMap[$hash].Add($fi)
        }
    }

    Write-Progress -Activity 'Step 2 - Hashing files' -Completed

    # Persist updated cache
    Save-HashCacheTable -Cache $hashCache -CacheFile $cacheFile
    Write-Host "Hash cache : $cacheHits hits / $cacheMisses computed — $cacheFile" -ForegroundColor DarkGray

    # Filter to groups with 2+ files (true duplicates)
    $dupGroups = $hashMap.GetEnumerator() |
        Where-Object { $_.Value.Count -gt 1 } |
        Select-Object @{N='Hash'; E={$_.Key}}, @{N='Files'; E={$_.Value}} |
        Sort-Object { ($_.Files | Measure-Object -Property Length -Sum).Sum } -Descending

    if (($dupGroups | Measure-Object).Count -eq 0) {
        Write-Host 'No duplicate files found. Your workspace is clean!' -ForegroundColor Green
        return
    }

    # ── Build resolution plan (default: keep newest) ──────────────────────────
    $resolvedGroups = [System.Collections.Generic.List[PSObject]]::new()
    $groupIdx       = 1

    foreach ($grp in $dupGroups) {
        $sorted   = $grp.Files | Sort-Object LastWriteTime -Descending
        $keepFile = $sorted[0].FullName   # newest by default

        $resolvedGroups.Add([PSCustomObject]@{
            GroupId    = "G$groupIdx"
            Hash       = $grp.Hash
            Files      = $grp.Files
            KeepFile   = $keepFile
            WasteBytes = ($grp.Files | Select-Object -Skip 1 | Measure-Object -Property Length -Sum).Sum
            Resolution = 'K'   # K=KeepNewest, O=KeepOldest, S=Skip
        })
        $groupIdx++
    }

    $totalWaste = ($resolvedGroups | Measure-Object -Property WasteBytes -Sum).Sum

    Write-Host ''
    Write-Host "Found $($resolvedGroups.Count) duplicate groups — $(Format-Size $totalWaste) reclaimable." -ForegroundColor Yellow
    Write-Host ''

    # ── Per-group resolution prompt ───────────────────────────────────────────
    Write-Host 'For each duplicate group, choose how to resolve it:' -ForegroundColor Cyan
    Write-Host '  K = Keep NEWEST  (delete older copies)  [default]' -ForegroundColor Gray
    Write-Host '  O = Keep OLDEST  (delete newer copies)' -ForegroundColor Gray
    Write-Host '  S = Skip         (keep all, delete none)' -ForegroundColor Gray
    Write-Host '  A = Apply K to ALL remaining groups' -ForegroundColor Gray
    Write-Host ''

    $applyAll = $false

    for ($i = 0; $i -lt $resolvedGroups.Count; $i++) {
        $g      = $resolvedGroups[$i]
        $sorted = $g.Files | Sort-Object LastWriteTime -Descending

        Write-Host "── Group $($g.GroupId) ($($g.Files.Count) copies, $(Format-Size $g.WasteBytes) wasted) ──" -ForegroundColor Yellow
        $fi = 0
        foreach ($f in $sorted) {
            $tag = if ($fi -eq 0) { ' [newest]' } else { ' [older]' }
            if ($fi -eq ($sorted.Count - 1) -and $fi -gt 0) { $tag = ' [oldest]' }
            Write-Host "  $($f.FullName)$tag" -ForegroundColor Gray
            $fi++
        }

        if ($applyAll) {
            Write-Host "  → Applying K (keep newest) automatically." -ForegroundColor DarkGray
            continue
        }

        $ans = (Read-Host "  Action? (K/O/S/A) [K]").ToUpper().Trim()
        if ([string]::IsNullOrWhiteSpace($ans)) { $ans = 'K' }

        switch ($ans) {
            'A' {
                $applyAll = $true
                $resolvedGroups[$i].Resolution = 'K'
                $resolvedGroups[$i].KeepFile   = ($sorted | Select-Object -First 1).FullName
            }
            'O' {
                $resolvedGroups[$i].Resolution = 'O'
                $resolvedGroups[$i].KeepFile   = ($sorted | Select-Object -Last 1).FullName
            }
            'S' {
                $resolvedGroups[$i].Resolution = 'S'
                $resolvedGroups[$i].KeepFile   = $null
            }
            default {
                $resolvedGroups[$i].Resolution = 'K'
                $resolvedGroups[$i].KeepFile   = ($sorted | Select-Object -First 1).FullName
            }
        }
        Write-Host ''
    }

    # Recalculate waste after resolution choices
    $actionableGroups = $resolvedGroups | Where-Object { $_.Resolution -ne 'S' }
    $finalWaste = 0L
    foreach ($g in $actionableGroups) {
        $sum = ($g.Files |
            Where-Object { $_.FullName -ne $g.KeepFile } |
            Measure-Object -Property Length -Sum).Sum
        $finalWaste += [long]$(if ($null -eq $sum) { 0 } else { $sum })
    }

    Write-Host ''
    Write-Host "Resolution summary:" -ForegroundColor Cyan
    Write-Host "  Groups to clean : $(($actionableGroups | Measure-Object).Count)" -ForegroundColor Yellow
    Write-Host "  Groups skipped  : $(($resolvedGroups | Where-Object { $_.Resolution -eq 'S' } | Measure-Object).Count)" -ForegroundColor Gray
    Write-Host "  Space to reclaim: $(Format-Size $finalWaste)" -ForegroundColor Yellow
    Write-Host ''

    # ── Generate HTML report ──────────────────────────────────────────────────
    $reportDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path $reportDir)) { New-Item -ItemType Directory -Path $reportDir | Out-Null }

    $ts         = Get-Date -Format 'yyyyMMdd-HHmmss'
    $reportPath = Join-Path $reportDir "Step2-Duplicates-$ts.html"

    New-Step2HtmlReport `
        -SearchedPaths   ($validPaths -join ', ') `
        -DuplicateGroups $resolvedGroups `
        -WasteBytes      $finalWaste `
        -OutputPath      $reportPath

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

    if (($actionableGroups | Measure-Object).Count -eq 0) {
        Write-Host 'All groups were skipped. Nothing to delete.' -ForegroundColor Cyan
        return
    }

    # ── Confirm deletion ──────────────────────────────────────────────────────
    Write-Host 'Type  DELETE  to confirm removal of duplicate copies, or press ENTER to cancel:' -ForegroundColor Red
    $confirm = Read-Host '  Confirm'

    if ($confirm -cne 'DELETE') {
        Write-Host 'Deletion cancelled. No files were removed.' -ForegroundColor Cyan
        return
    }

    # ── Perform deletion ──────────────────────────────────────────────────────
    $logPath = Join-Path $reportDir "Step2-DeletionLog-$ts.txt"
    "WinCleaner Step 2 — Deletion Log"       | Set-Content  $logPath -Encoding UTF8
    "Date   : $(Get-Date)"                   | Add-Content  $logPath
    "Paths  : $($validPaths -join ', ')"     | Add-Content  $logPath
    ('─' * 60)                               | Add-Content  $logPath

    $deleted = 0
    $errored = 0

    foreach ($g in $actionableGroups) {
        $toDelete = $g.Files | Where-Object { $_.FullName -ne $g.KeepFile }
        "[GROUP $($g.GroupId) — keep: $($g.KeepFile)]" | Add-Content $logPath

        foreach ($f in $toDelete) {
            if (-not (Test-Path -LiteralPath $f.FullName)) { continue }
            try {
                Remove-Item -LiteralPath $f.FullName -Force
                "[DELETED] $($f.FullName)" | Add-Content $logPath
                $deleted++
                Write-Host "  Deleted: $($f.FullName)" -ForegroundColor DarkGray
            }
            catch {
                "[ERROR]   $($f.FullName) — $($_.Exception.Message)" | Add-Content $logPath
                $errored++
                Write-Host "  ERROR  : $($f.FullName)" -ForegroundColor Red
            }
        }
    }

    Write-Host ''
    Write-Host '── Step 2 Complete ───────────────────────────────────────────' -ForegroundColor Green
    Write-Host "   Deleted : $deleted duplicate files"                          -ForegroundColor Green
    if ($errored -gt 0) {
        Write-Host "   Errors  : $errored files (check log)"                    -ForegroundColor Yellow
    }
    Write-Host "   Log     : $logPath"                                          -ForegroundColor Gray
    Write-Host ''
}

# ── Entry point ────────────────────────────────────────────────────────────────
if ($MyInvocation.InvocationName -ne '.') {
    Invoke-Step2 -Paths $SearchPaths -MinKB $MinSizeKB -DryRunMode $DryRun.IsPresent
}
