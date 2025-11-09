param(
  [string]$ExcelPath = 'C:\AI\asagake\ASAGAKE.xlsm',
  [string]$Repo = 'C:\AI\asagake',
  [int]$IntervalSec = 5,
  [int]$RetainDays = 30
)

$TaskName = 'ASAGAKE_BoardLogger'
$Python = 'C:\\Python313\\python.exe'
$Action = New-ScheduledTaskAction -Execute $Python -Argument "scripts/board_logger_daemon.py --dashboard $ExcelPath --dash-outdir $Repo\output\j_logs --interval $IntervalSec --retain-days $RetainDays" -WorkingDirectory $Repo

# 09:00-15:35 JST (平日)に5分間隔で繰返し起動し、同名デーモンは自己制御
$Trigger = New-ScheduledTaskTrigger -Daily -At 09:00
$Trigger.Repetition = New-ScheduledTaskRepetitionSettings -Interval (New-TimeSpan -Minutes 5) -Duration (New-TimeSpan -Hours 6 -Minutes 35)

try {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
} catch {}

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Description 'Capture J logs during market hours' -User "$env:USERNAME" -RunLevel Highest | Out-Null
Write-Output "registered=$TaskName"

