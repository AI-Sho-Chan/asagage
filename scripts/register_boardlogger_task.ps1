param(
  [string]$ExcelPath = 'C:\AI\asagake\ASAGAKE.xlsm',
  [string]$Repo = 'C:\AI\asagake',
  [int]$IntervalSec = 5,
  [int]$RetainDays = 30
)

$TaskName = 'ASAGAKE_BoardLogger'
$Python = 'C:\Python313\python.exe'
$TaskCmd = Join-Path $Repo 'scripts\board_logger_task.cmd'
$TaskCmdQuoted = '"' + $TaskCmd + '"'
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
  # 09:00~15:35 の間に5分おきで起動するタスクを登録（schtasks版）
  # 既存の CMD が壊れているケースを避けるため、毎回内容を上書き生成する
  Set-Content -Path $TaskCmd -Value "@echo off`r`ncd /d $Repo`r`n$Cmd`r`n" -Encoding ASCII
  try { schtasks.exe /Delete /F /TN $TaskName | Out-Null } catch {}
  $args = @('/Create','/F','/TN',$TaskName,'/SC','DAILY','/ST','09:00','/DU','06:35','/RI','5','/TR',$TaskCmdQuoted)
  schtasks.exe @args | Out-Null
}

try {
  $used = ''
  if (Get-Command New-ScheduledTaskAction -ErrorAction SilentlyContinue) {
    try {
      Register-WithScheduledTaskModule
      $used = 'ScheduledTask'
    } catch {
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
