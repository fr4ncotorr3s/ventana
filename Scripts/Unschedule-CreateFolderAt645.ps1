$taskName = "CreateFolderAt645Task"

$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

if ($task) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "Scheduled task '$taskName' has been successfully removed."
}
else {
    Write-Host "Scheduled task '$taskName' does not exist."
}
