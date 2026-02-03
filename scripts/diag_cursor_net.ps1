<#
.SYNOPSIS
  Diagnose Cursor network/DNS issues and write evidence logs.

.DESCRIPTION
  Writes a timestamped append-only log under: C:\AI\asagake\logs\
  This is designed to help explain "turn_aborted" incidents without blaming user actions.

  Safe by default:
    - No changes are made unless you pass switches such as -AutoFix / -RestartCursor / -SetDnsDiversified.

  Notes:
    - Restarting Cursor may lose unsaved work. It is OFF by default.
    - Changing DNS requires Administrator privileges. It is OFF by default.
#>

[CmdletBinding()]
param(
  [string]$HostName = "api2.cursor.sh",
  [string]$RepoRoot = "",
  # Optional: explicitly point to Cursor logs root, useful when running as SYSTEM.
  # Example: C:\Users\PC_User\AppData\Roaming\Cursor\logs
  [string]$CursorLogsRoot = "",
  [switch]$AutoFix,
  [switch]$RestartCursor,
  [switch]$SetDnsDiversified,
  [string]$InterfaceAlias = ""
)

Set-StrictMode -Version Latest

function Test-IsAdmin {
  $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = New-Object Security.Principal.WindowsPrincipal($identity)
  return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
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

function Try-Run {
  param(
    [string]$LogPath,
    [string]$Title,
    [scriptblock]$Action
  )

  Write-LogLine -Path $LogPath -Line "=== $Title ==="
  try {
    $out = & $Action
    if ($null -ne $out) {
      ($out | Out-String).TrimEnd("`r","`n").Split("`n") | ForEach-Object {
        $line = $_.TrimEnd("`r")
        if ($line) { Write-LogLine -Path $LogPath -Line $line }
      }
    }
  } catch {
    Write-LogLine -Path $LogPath -Line "[error] $($_.Exception.GetType().Name): $($_.Exception.Message)"
  }
}

function Detect-PrimaryInterfaceAlias {
  # Prefer the interface used by the default route (0.0.0.0/0).
  try {
    $route = Get-NetRoute -DestinationPrefix "0.0.0.0/0" -ErrorAction Stop |
      Sort-Object -Property RouteMetric, InterfaceMetric |
      Select-Object -First 1
    if ($route -and $route.InterfaceAlias) {
      return [string]$route.InterfaceAlias
    }
  } catch {}

  # Fallback: the first interface that has IPv4 DNS servers configured.
  try {
    $dns = Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction Stop |
      Where-Object { $_.ServerAddresses -and $_.ServerAddresses.Count -gt 0 } |
      Select-Object -First 1
    if ($dns -and $dns.InterfaceAlias) {
      return [string]$dns.InterfaceAlias
    }
  } catch {}

  return ""
}

function Scan-CursorLogs {
  param(
    [string]$LogPath,
    [string]$HostForPattern,
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

  if (-not $cursorLogRoot -or -not (Test-Path -LiteralPath $cursorLogRoot)) {
    $shown = $LogsRoot
    if (-not $shown) { $shown = (Join-Path $env:APPDATA "Cursor\\logs") }
    Write-LogLine -Path $LogPath -Line "[warn] Cursor logs not found: $shown"
    return
  }

  $pattern = "ENOTFOUND\\s+$([Regex]::Escape($HostForPattern))|EAI_AGAIN|ECONNRESET|ETIMEDOUT|socket hang up"
  $dirs = Get-ChildItem -LiteralPath $cursorLogRoot -Directory -ErrorAction SilentlyContinue |
    Sort-Object Name -Descending |
    Select-Object -First 10

  foreach ($d in $dirs) {
    $files = @(
      (Join-Path $d.FullName "window1\\exthost\\exthost.log"),
      (Join-Path $d.FullName "window1\\renderer.log")
    ) | Where-Object { Test-Path -LiteralPath $_ }

    foreach ($f in $files) {
      $matches = Select-String -LiteralPath $f -Pattern $pattern -AllMatches -ErrorAction SilentlyContinue
      if (-not $matches) { continue }

      $matchList = @($matches)
      Write-LogLine -Path $LogPath -Line "[CursorLog] $f hits=$($matchList.Count)"
      $tail = $matchList | Select-Object -Last 3
      foreach ($m in $tail) {
        $line = $m.Line
        if ($line.Length -gt 350) { $line = $line.Substring(0, 350) + " ..." }
        Write-LogLine -Path $LogPath -Line ("[CursorLog] " + $line)
      }
    }
  }
}

$repo = Get-RepoRootFromScript
$logDir = Join-Path $repo "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

$stamp = Get-Date -Format "yyyyMMdd"
$logFile = Join-Path $logDir ("cursor_netdiag_{0}.log" -f $stamp)

Write-LogLine -Path $logFile -Line "===== Cursor NetDiag start ====="
Write-LogLine -Path $logFile -Line "repo=$repo"
Write-LogLine -Path $logFile -Line "host=$HostName"
Write-LogLine -Path $logFile -Line ("whoami=" + (& whoami))
Write-LogLine -Path $logFile -Line ("is_admin=" + (Test-IsAdmin))
Write-LogLine -Path $logFile -Line ("computer=" + $env:COMPUTERNAME)

Try-Run -LogPath $logFile -Title "DNS servers (IPv4)" -Action {
  Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction Stop |
    Select-Object InterfaceAlias, ServerAddresses
}

Try-Run -LogPath $logFile -Title "Resolve-DnsName (default)" -Action {
  Resolve-DnsName -Name $HostName -Type A -ErrorAction Stop | Select-Object Name, IPAddress
}
Try-Run -LogPath $logFile -Title "Resolve-DnsName (1.1.1.1)" -Action {
  Resolve-DnsName -Name $HostName -Type A -Server 1.1.1.1 -ErrorAction Stop | Select-Object Name, IPAddress
}
Try-Run -LogPath $logFile -Title "Resolve-DnsName (8.8.8.8)" -Action {
  Resolve-DnsName -Name $HostName -Type A -Server 8.8.8.8 -ErrorAction Stop | Select-Object Name, IPAddress
}

Try-Run -LogPath $logFile -Title "Test-NetConnection 443" -Action {
  Test-NetConnection -ComputerName $HostName -Port 443 -InformationLevel Detailed -ErrorAction Stop |
    Select-Object ComputerName, RemoteAddress, RemotePort, TcpTestSucceeded, PingSucceeded
}

Try-Run -LogPath $logFile -Title "Default route" -Action {
  Get-NetRoute -DestinationPrefix "0.0.0.0/0" -ErrorAction Stop |
    Sort-Object -Property RouteMetric, InterfaceMetric |
    Select-Object -First 3
}

Scan-CursorLogs -LogPath $logFile -HostForPattern $HostName -LogsRoot $CursorLogsRoot

if ($AutoFix) {
  Try-Run -LogPath $logFile -Title "AutoFix: ipconfig /flushdns" -Action {
    & ipconfig /flushdns
  }
}

if ($SetDnsDiversified) {
  if (-not (Test-IsAdmin)) {
    Write-LogLine -Path $logFile -Line "[warn] -SetDnsDiversified requires Administrator. Skipped."
  } else {
    if (-not $InterfaceAlias) {
      $InterfaceAlias = Detect-PrimaryInterfaceAlias
    }
    if (-not $InterfaceAlias) {
      Write-LogLine -Path $logFile -Line "[warn] Could not detect InterfaceAlias. Use -InterfaceAlias explicitly."
    } else {
      Try-Run -LogPath $logFile -Title "Set DNS servers: 1.1.1.1, 8.8.8.8 (InterfaceAlias=$InterfaceAlias)" -Action {
        Set-DnsClientServerAddress -InterfaceAlias $InterfaceAlias -ServerAddresses @("1.1.1.1", "8.8.8.8") -ErrorAction Stop
        Get-DnsClientServerAddress -InterfaceAlias $InterfaceAlias -AddressFamily IPv4 | Select-Object InterfaceAlias, ServerAddresses
      }
      Try-Run -LogPath $logFile -Title "ipconfig /flushdns (after DNS change)" -Action {
        & ipconfig /flushdns
      }
    }
  }
}

if ($RestartCursor) {
  Write-LogLine -Path $logFile -Line "[warn] RestartCursor requested. Unsaved work may be lost."
  Try-Run -LogPath $logFile -Title "Stop Cursor process (if any)" -Action {
    Get-Process -Name "Cursor" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    "stopped"
  }

  Try-Run -LogPath $logFile -Title "Start Cursor (best-effort)" -Action {
    $cmd = Get-Command -Name "cursor" -ErrorAction SilentlyContinue
    if ($cmd) {
      Start-Process -FilePath $cmd.Source | Out-Null
      "started via 'cursor' command"
      return
    }

    $defaultPath = Join-Path $env:LOCALAPPDATA "Programs\\Cursor\\Cursor.exe"
    if (Test-Path -LiteralPath $defaultPath) {
      Start-Process -FilePath $defaultPath | Out-Null
      "started via $defaultPath"
      return
    }

    "[warn] Cursor executable not found; please start Cursor manually."
  }
}

Write-LogLine -Path $logFile -Line "===== Cursor NetDiag end ====="
Write-Output $logFile
