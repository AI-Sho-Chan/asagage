<#
.SYNOPSIS
  Watch Cursor logs for ECONNRESET and write an evidence log (one-shot).

.DESCRIPTION
  This script is designed as a lightweight "alarm" that can be run on a schedule.
  It reads only *new bytes* from the latest Cursor log files and detects ECONNRESET.

  Step-1 (ALERT) version: detection + logging only (no auto-fix / no Cursor restart).
  In the next step we can wire this to call scripts/diag_cursor_net.ps1 automatically.

.OUTPUTS
  - Appends to: C:\AI\asagake\logs\cursor_econnreset_watch_YYYYMMDD.log
  - Persists offsets to: C:\AI\asagake\logs\cursor_econnreset_watch_state.json

.EXITCODES
  0 = no new ECONNRESET detected
  10 = ECONNRESET detected (new since last run)
  2 = script error (still tries to persist state)
#>

[CmdletBinding()]
param(
  [string]$RepoRoot = "",
  # NOTE: Task Scheduler / callers can sometimes pass this as an array-like value.
  # Keep it as [object] and normalize to int internally to avoid hard failures.
  [object]$MaxDirs = 5,
  # Optional: explicitly point to Cursor logs root, useful when running as SYSTEM.
  # Example: C:\Users\PC_User\AppData\Roaming\Cursor\logs
  [string]$CursorLogsRoot = "",

  # Step-2: If enabled, automatically run scripts/diag_cursor_net.ps1 when ECONNRESET is detected.
  [switch]$RunDiagOnHit,
  [int]$DiagCooldownMinutes = 10,
  [string]$DiagHostName = "api2.cursor.sh"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Normalize-PositiveInt {
  param(
    [object]$Value,
    [int]$DefaultValue
  )

  try {
    if ($null -eq $Value) { return $DefaultValue }
    $first = @($Value)[0]
    if ($null -eq $first) { return $DefaultValue }
    $n = [int]$first
    if ($n -le 0) { return $DefaultValue }
    return $n
  } catch {
    return $DefaultValue
  }
}

function Get-RepoRootFromScript {
  if ($RepoRoot -and (Test-Path -LiteralPath $RepoRoot)) {
    return (Resolve-Path -LiteralPath $RepoRoot).Path
  }
  if ($PSScriptRoot) {
    $candidate = Join-Path $PSScriptRoot ".."
    if (Test-Path -LiteralPath $candidate) {
      return (Resolve-Path -LiteralPath $candidate).Path
    }
  }
  return (Get-Location).Path
}

function Write-LogLine {
  param(
    [string]$Path,
    [string]$Line
  )
  $timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffK")
  Add-Content -LiteralPath $Path -Value "$timestamp $Line" -Encoding utf8
}

function Load-State {
  param([string]$Path)
  $default = @{
    version = 1
    file_offsets = @{}
    last_scan_utc = ""
    last_diag_utc = ""
  }

  if (-not (Test-Path -LiteralPath $Path)) {
    return $default
  }

  try {
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if (-not $raw) { return $default }
    $obj = $raw | ConvertFrom-Json -ErrorAction Stop

    # NOTE: StrictMode=Latest makes "missing property" a hard error on PSCustomObject.
    # We intentionally treat missing fields as default values (backward-compatible).
    $props = @()
    try { $props = $obj.PSObject.Properties.Name } catch { $props = @() }

    $offsets = @{}
    if (($props -contains "file_offsets") -and $obj.file_offsets) {
      $obj.file_offsets.psobject.Properties | ForEach-Object {
        $k = [string]$_.Name
        if ($k) {
          $kResolved = $k
          try { $kResolved = (Resolve-Path -LiteralPath $k -ErrorAction Stop).Path } catch { $kResolved = $k }
          $offsets[$kResolved.ToLowerInvariant()] = [long]$_.Value
        }
      }
    }

    $version = 1
    if (($props -contains "version") -and $obj.version) { $version = [int]$obj.version }

    $lastScan = ""
    if (($props -contains "last_scan_utc") -and $obj.last_scan_utc) { $lastScan = [string]$obj.last_scan_utc }

    $lastDiag = ""
    if (($props -contains "last_diag_utc") -and $obj.last_diag_utc) { $lastDiag = [string]$obj.last_diag_utc }

    return @{
      version = $version
      file_offsets = $offsets
      last_scan_utc = $lastScan
      last_diag_utc = $lastDiag
    }
  } catch {
    $default.state_load_error = $_.Exception.Message
    return $default
  }
}

function Should-RunDiagNow {
  param(
    [object]$State,
    [int]$CooldownMinutes
  )
  try {
    if (-not $CooldownMinutes -or $CooldownMinutes -le 0) { return $true }
    $lastStr = ""
    try { $lastStr = [string]$State.last_diag_utc } catch { $lastStr = "" }
    if (-not $lastStr) { return $true }

    $last = [DateTime]::Parse($lastStr, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
    $now = (Get-Date).ToUniversalTime()
    return (($now - $last).TotalMinutes -ge $CooldownMinutes)
  } catch {
    return $true
  }
}

function Run-DiagOnHit {
  param(
    [string]$Repo,
    [string]$LogPath,
    [string]$HostName,
    [string]$LogsRoot
  )

  $diagScript = Join-Path $Repo "scripts\\diag_cursor_net.ps1"
  if (-not (Test-Path -LiteralPath $diagScript)) {
    Write-LogLine -Path $LogPath -Line "[warn] diag script not found: $diagScript"
    return
  }

  try {
    $start = Get-Date
    Write-LogLine -Path $LogPath -Line "[diag] start host=$HostName script=$diagScript"
    if ($LogsRoot) {
      $out = & $diagScript -HostName $HostName -RepoRoot $Repo -CursorLogsRoot $LogsRoot 2>&1
    } else {
      $out = & $diagScript -HostName $HostName -RepoRoot $Repo 2>&1
    }
    $elapsed = (New-TimeSpan -Start $start -End (Get-Date)).TotalSeconds
    $outStr = ($out | Out-String).Trim()
    if ($outStr) {
      $lines = $outStr -split "`r?`n"
      foreach ($l in ($lines | Select-Object -Last 3)) {
        Write-LogLine -Path $LogPath -Line ("[diag] " + $l.TrimEnd("`r"))
      }
    }
    Write-LogLine -Path $LogPath -Line ("[diag] end elapsed_sec={0:N1}" -f $elapsed)
  } catch {
    Write-LogLine -Path $LogPath -Line "[diag][error] $($_.Exception.GetType().Name): $($_.Exception.Message)"
  }
}

function Save-State {
  param(
    [string]$Path,
    [object]$State
  )
  $dir = Split-Path -Parent $Path
  New-Item -ItemType Directory -Force -Path $dir | Out-Null

  $tmp = "$Path.tmp"
  ($State | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $tmp -Encoding utf8
  Move-Item -LiteralPath $tmp -Destination $Path -Force
}

function Read-NewText {
  param(
    [string]$Path,
    [long]$Offset
  )
  if (-not (Test-Path -LiteralPath $Path)) { return @{ text = ""; offset = 0; exists = $false; len = 0 } }

  $fi = Get-Item -LiteralPath $Path -ErrorAction Stop
  $len = [long]$fi.Length
  if ($Offset -gt $len) { $Offset = 0 } # truncated / rotated

  $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
  try {
    [void]$fs.Seek($Offset, [System.IO.SeekOrigin]::Begin)
    $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true)
    $text = $sr.ReadToEnd()
    $newOffset = [long]$fs.Position
    return @{ text = $text; offset = $newOffset; exists = $true; len = $len }
  } finally {
    $fs.Dispose()
  }
}

function Get-Latest-CursorLogFiles {
  param(
    [int]$DirLimit,
    [string]$LogsRoot
  )

  $cursorLogRoot = ""
  if ($LogsRoot) {
    $cursorLogRoot = $LogsRoot
  } else {
    $cursorLogRoot = Join-Path $env:APPDATA "Cursor\\logs"
    if (-not (Test-Path -LiteralPath $cursorLogRoot)) {
      # Fallback: try to find a plausible user profile Cursor logs folder (for SYSTEM runs).
      try {
        $candidates = Get-ChildItem -LiteralPath "C:\\Users" -Directory -ErrorAction SilentlyContinue |
          ForEach-Object { Join-Path $_.FullName "AppData\\Roaming\\Cursor\\logs" } |
          Where-Object { Test-Path -LiteralPath $_ }

        if ($candidates -and $candidates.Count -gt 0) {
          # Pick the one with the newest subdirectory name (Cursor log dirs are timestamp-like).
          $best = $null
          $bestScore = ""
          foreach ($c in $candidates) {
            $d = Get-ChildItem -LiteralPath $c -Directory -ErrorAction SilentlyContinue |
              Sort-Object Name -Descending |
              Select-Object -First 1
            if ($d -and $d.Name) {
              if (-not $best -or $d.Name -gt $bestScore) {
                $best = $c
                $bestScore = $d.Name
              }
            }
          }
          if ($best) { $cursorLogRoot = $best }
        }
      } catch {}
    }
  }

  if (-not $cursorLogRoot -or -not (Test-Path -LiteralPath $cursorLogRoot)) { return @() }

  $dirList = Get-ChildItem -LiteralPath $cursorLogRoot -Directory -ErrorAction SilentlyContinue |
    Sort-Object Name -Descending |
    Select-Object -First $DirLimit

  $files = @()
  foreach ($d in $dirList) {
    $candidates = @(
      (Join-Path $d.FullName "window1\\exthost\\exthost.log"),
      (Join-Path $d.FullName "window1\\renderer.log")
    )
    foreach ($f in $candidates) {
      if (Test-Path -LiteralPath $f) { $files += $f }
    }
  }
  return ($files | Select-Object -Unique)
}

$repo = Get-RepoRootFromScript
$logDir = Join-Path $repo "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

$dateTag = Get-Date -Format "yyyyMMdd"
$logPath = Join-Path $logDir ("cursor_econnreset_watch_{0}.log" -f $dateTag)
$statePath = Join-Path $logDir "cursor_econnreset_watch_state.json"

$exitCode = 0
$state = Load-State -Path $statePath

try {
  $dirLimit = Normalize-PositiveInt -Value $MaxDirs -DefaultValue 5
  $files = Get-Latest-CursorLogFiles -DirLimit $dirLimit -LogsRoot $CursorLogsRoot
  if (-not $files -or $files.Count -eq 0) {
    $rootShown = $CursorLogsRoot
    if (-not $rootShown) { $rootShown = (Join-Path $env:APPDATA "Cursor\\logs") }
    Write-LogLine -Path $logPath -Line "[warn] Cursor log files not found under $rootShown"
    $exitCode = 2
  } else {
    $found = $false
    foreach ($f in $files) {
      $canonical = $f
      try { $canonical = (Resolve-Path -LiteralPath $f -ErrorAction Stop).Path } catch { $canonical = $f }
      $key = ([string]$canonical).ToLowerInvariant()
      $offset = 0
      if ($state.file_offsets.ContainsKey($key)) {
        $offset = [long]$state.file_offsets[$key]
      } else {
        # First time seeing this file: baseline to EOF to avoid firing on historical errors.
        try {
          $fi = Get-Item -LiteralPath $f -ErrorAction Stop
          $state.file_offsets[$key] = [long]$fi.Length
          Write-LogLine -Path $logPath -Line "[init] baseline offset=$($fi.Length) file=$f"
          continue
        } catch {}
      }
      $res = Read-NewText -Path $f -Offset $offset
      $state.file_offsets[$key] = [long]$res.offset
      if (-not $res.exists) { continue }

      if ($res.text -and $res.text -match "ECONNRESET") {
        $found = $true
        Write-LogLine -Path $logPath -Line "[hit] ECONNRESET detected in $f (len=$($res.len), from_offset=$offset -> $($res.offset))"
        $lines = $res.text -split "`r?`n"
        $hits = $lines | Where-Object { $_ -match "ECONNRESET" } | Select-Object -Last 3
        foreach ($h in $hits) {
          $line = $h
          if ($line.Length -gt 350) { $line = $line.Substring(0, 350) + " ..." }
          Write-LogLine -Path $logPath -Line ("[hit] " + $line)
        }
      }
    }

    if ($found) {
      $exitCode = 10
      if ($RunDiagOnHit) {
        if (Should-RunDiagNow -State $state -CooldownMinutes $DiagCooldownMinutes) {
          Run-DiagOnHit -Repo $repo -LogPath $logPath -HostName $DiagHostName -LogsRoot $CursorLogsRoot
          try { $state.last_diag_utc = (Get-Date).ToUniversalTime().ToString("o") } catch {}
        } else {
          Write-LogLine -Path $logPath -Line ("[diag] skipped (cooldown {0} min)" -f $DiagCooldownMinutes)
        }
      }
    }
  }
} catch {
  $exitCode = 2
  Write-LogLine -Path $logPath -Line "[error] $($_.Exception.GetType().Name): $($_.Exception.Message)"
} finally {
  $state.last_scan_utc = (Get-Date).ToUniversalTime().ToString("o")
  Save-State -Path $statePath -State $state
}

exit $exitCode
