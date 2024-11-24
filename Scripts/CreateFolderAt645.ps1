$desktopPath = [System.IO.Path]::Combine($env:USERPROFILE, 'Desktop')
$documentsPath = [System.IO.Path]::Combine($env:USERPROFILE, 'Documents')

$targetTime = Get-Date -Hour 6 -Minute 45 -Second 0

if ((Get-Date) -lt $targetTime) {
    $targetTime = $targetTime.AddDays(1)
}

$folderName = $targetTime.ToString('yyyy-MM-dd_HH-mm-ss')
$folderPath = [System.IO.Path]::Combine($desktopPath, $folderName)

function Create-Shortcut ($folderPath) {
    $WScriptShell = New-Object -ComObject WScript.Shell
    $shortcutPath = [System.IO.Path]::Combine($desktopPath, "$folderName.lnk")
    $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $folderPath
    $shortcut.Save()
    return $shortcutPath
}

function PinToStart($shortcutPath) {
    try {
        $shell = New-Object -ComObject shell.application
        $folder = $shell.Namespace([System.IO.Path]::GetDirectoryName($shortcutPath))
        $item = $folder.ParseName([System.IO.Path]::GetFileName($shortcutPath))
        
        $verb = $item.Verbs() | Where-Object { $_.Name -eq 'Pin to Start Men&u' }
        if ($verb) {
            $verb.DoIt()
            Write-Host "Pinned to Start: $shortcutPath"
        }
        else {
            Write-Host "Pin to Start option not found. Trying fallback method."
            
            $startMenuPath = [System.IO.Path]::Combine($desktopPath, "$folderName.lnk")
            $startPinVerb = $item.Verbs() | Where-Object { $_.Name -eq 'Pin to Start' }
            if ($startPinVerb) {
                $startPinVerb.DoIt()
                Write-Host "Pinned to Start via fallback method: $shortcutPath"
            }
            else {
                Write-Host "Pin to Start still not available in fallback."
            }
        }
    }
    catch {
        Write-Host "Error pinning to Start: $_"
    }
}

function PinToQuickAccess($shortcutPath) {
    try {
        $shell = New-Object -ComObject shell.application
        $folder = $shell.Namespace([System.IO.Path]::GetDirectoryName($shortcutPath))
        $item = $folder.ParseName([System.IO.Path]::GetFileName($shortcutPath))
        
        $verb = $item.Verbs() | Where-Object { $_.Name -eq 'Pin to Quick access' }
        if ($verb) {
            $verb.DoIt()
            Write-Host "Pinned to Quick Access: $shortcutPath"
        }
        else {
            Write-Host "Pin to Quick Access option not found."
        }
    }
    catch {
        Write-Host "Error pinning to Quick Access: $_"
    }
}

if ((Get-Date) -gt $targetTime) {
    if (-not (Test-Path $folderPath)) {
        New-Item -Path $folderPath -ItemType Directory
        Write-Host "Folder created: $folderPath"
        
        $shortcutPath = Create-Shortcut $folderPath
        
        PinToStart $shortcutPath
        PinToQuickAccess $shortcutPath
    }
    else {
        Write-Host "Folder already exists, skipping creation."
    }
}

$timeNow = Get-Date
$timeToWait = (Get-Date).Date.AddHours(20) - $timeNow
if ($timeToWait.TotalSeconds -gt 0) {
    Write-Host "Waiting until 8:00 PM..."
    Start-Sleep -Seconds $timeToWait.TotalSeconds
}

if (Test-Path $folderPath) {
    $destinationPath = [System.IO.Path]::Combine($documentsPath, $folderName)
    Copy-Item -Path $folderPath -Destination $destinationPath -Recurse
    Write-Host "Folder copied to Documents: $destinationPath"

    $shell = New-Object -ComObject shell.application
    $folder = $shell.Namespace($folderPath)
    if ($folder) {
        $folderItem = $folder.Items() | Where-Object { $_.Name -eq (Get-Item $folderPath).Name }
        if ($folderItem) {
            $folderItem.InvokeVerb('Unpin from Start')
            $folderItem.InvokeVerb('Unpin from Quick access')
            Write-Host "Unpinned from Start and Quick Access."
        }
        else {
            Write-Host "Folder item not found to unpin."
        }
    }
    else {
        Write-Host "Failed to access the folder for unpinning."
    }

    if (Test-Path $destinationPath) {
        Remove-Item -Path $folderPath -Recurse
        Write-Host "Folder removed from Desktop."
    }
}
