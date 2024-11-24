$scriptPath = ".\CreateFolderAt645.ps1"

$taskName = "CreateFolderAt645Task"
$taskDescription = "This task runs a PowerShell script to create a folder at 6:45 AM and handle related tasks."

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$scriptPath`""

$trigger = New-ScheduledTaskTrigger -AtStartup

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -StartWhenAvailable -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5)

$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal -TaskName $taskName -Description $taskDescription

Write-Host "Scheduled task created successfully."
