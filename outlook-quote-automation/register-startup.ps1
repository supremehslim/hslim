# Windows PowerShell에서 실행. 현재 사용자 로그인 시 자동화가 시작됩니다.
$ErrorActionPreference = 'Stop'
$taskName = 'OutlookQuoteDraftAutomation'
$pythonPath = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $pythonPath)) { throw 'setup.bat을 먼저 실행하세요.' }
$action = New-ScheduledTaskAction -Execute $pythonPath -Argument '-m quote_automation watch' -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -AtLogOn -User ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
$principal = New-ScheduledTaskPrincipal -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
Write-Host "등록 완료: $taskName. 다음 Windows 로그인부터 실행됩니다."
Write-Host "삭제: Unregister-ScheduledTask -TaskName $taskName -Confirm:`$false"
