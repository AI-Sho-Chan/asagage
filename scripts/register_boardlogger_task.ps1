param(
  [string]$ExcelPath = 'C:\AI\asagake\ASAGAKE.xlsm',
  [string]$Repo = 'C:\AI\asagake',
  [int]$IntervalSec = 5,
  [int]$RetainDays = 30
)

$TaskName = 'ASAGAKE_BoardLogger'
$Python = 'C:\\Python313\\python.exe'
$Cmd = "$Python scripts/board_logger_daemon.py --dashboard `"$ExcelPath`" --dash-outdir `"$Repo\output\j_logs`" --interval $IntervalSec --retain-days $RetainDays"
$CmdPs = "Set-Location `'$Repo`'; $Cmd"

function Register-WithScheduledTaskModule {
  param()
  $Action = New-ScheduledTaskAction -Execute $Python -Argument "scripts/board_logger_daemon.py --dashboard $ExcelPath --dash-outdir $Repo\output\j_logs --interval $IntervalSec --retain-days $RetainDays" -WorkingDirectory $Repo
  $Trigger = New-ScheduledTaskTrigger -Daily -At 09:00
  $Trigger.Repetition = New-ScheduledTaskRepetitionSettings -Interval (New-TimeSpan -Minutes 5) -Duration (New-TimeSpan -Hours 6 -Minutes 35)
  try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
  Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Description 'Capture J logs during market hours' -User "$env:USERNAME" -RunLevel Highest | Out-Null
}

function Register-WithSchTasksExe {
  param()
  # 09:00 から 6時間35分の間、5分間隔で起動
  $args = "/Create /F /TN `"$TaskName`" /SC MINUTE /MO 5 /ST 09:00 /DU 06:35 /TR `"powershell -NoProfile -ExecutionPolicy Bypass -Command `'$CmdPs`'`""
  try { schtasks.exe /Delete /F /TN $TaskName | Out-Null } catch {}
  schtasks.exe $args | Out-Null
}

try {
  $used = ''
  if (Get-Command New-ScheduledTaskAction -ErrorAction SilentlyContinue) {
    try {
      Register-WithScheduledTaskModule
      $used = 'ScheduledTask'
    } catch {
      # 一部環境で RepetitionSettings が未実装のためフォールバック
      Register-WithSchTasksExe
      $used = 'schtasks'
    }
  }
  else {
    Register-WithSchTasksExe
    $used = 'schtasks'
  }
  Write-Output "registered=$TaskName via=$used"
}
catch {
  Write-Output "register_error=$($_.Exception.Message)"
}
