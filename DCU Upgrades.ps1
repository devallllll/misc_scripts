<#
  .SYNOPSIS
    RMM script to run Dell Command Update
  .DESCRIPTION
    Removes Dell bloatware, runs Dell Command Update for all updates including BIOS,
    and creates post-reboot task if needed.
#>

function Remove-DellBloatware {
    $ProcessesToStop = @('SupportAssistClientUI', 'SupportAssistAgent', 'DellClientManagementService')
    foreach ($ProcessName in $ProcessesToStop) {
        Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    
    $AppsToRemove = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' | 
        Where-Object { $_.DisplayName -like '*SupportAssist*' -or $_.DisplayName -like '*Dell Update*' -or $_.DisplayName -like '*Dell Foundation Services*' }
    
    foreach ($App in $AppsToRemove) {
        try {
            if ($App.UninstallString -match 'msiexec') {
                $Guid = [regex]::Match($App.UninstallString, '\{[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}\}').Value
                Start-Process -NoNewWindow -Wait -FilePath 'msiexec.exe' -ArgumentList "/x $Guid /quiet /qn"
            }
            else { 
                Start-Process -NoNewWindow -Wait -FilePath $App.UninstallString -ArgumentList '/quiet' -ErrorAction SilentlyContinue
            }
        }
        catch { }
    }
}

function Create-PostRebootTask {
    $taskScript = @"
`$DCU = Get-ChildItem -Path "`$env:SystemDrive\Program Files*\Dell\CommandUpdate\dcu-cli.exe" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
if (`$DCU) {
    Start-Process -NoNewWindow -Wait -FilePath `$DCU -ArgumentList '/applyUpdates -silent'
    `$rebootCheck = Start-Process -NoNewWindow -Wait -FilePath `$DCU -ArgumentList '/rebootpending' -PassThru
    if (`$rebootCheck.ExitCode -ne 0) {
        Unregister-ScheduledTask -TaskName "DellUpdatePostReboot" -Confirm:`$false -ErrorAction SilentlyContinue
    }
}
"@
    
    $taskScriptPath = "$env:TEMP\DellUpdatePostReboot.ps1"
    $taskScript | Out-File -FilePath $taskScriptPath -Encoding UTF8
    
    $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -File `"$taskScriptPath`""
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    
    Register-ScheduledTask -TaskName "DellUpdatePostReboot" -Action $action -Trigger $trigger -Principal $principal -Description "Complete Dell updates after reboot" -Force | Out-Null
}

# Find DCU
$DCU = Get-ChildItem -Path "$env:SystemDrive\Program Files*\Dell\CommandUpdate\dcu-cli.exe" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName

if (-not $DCU) {
    Write-Output "Dell Command Update CLI not found."
    exit 1
}

Write-Output "Starting Dell update process..."

# Remove bloatware
Remove-DellBloatware

# Configure DCU for all updates including BIOS with BitLocker auto-suspend
Start-Process -NoNewWindow -Wait -FilePath $DCU -ArgumentList '/configure -autoSuspendBitLocker=enable -scheduleAction=DownloadInstallAndNotify -updatesNotification=disable -scheduleAuto -silent'

# Run updates in background - prompts user for reboot
Start-Process -FilePath $DCU -ArgumentList '/applyUpdates -autoSuspendBitLocker=enable -reboot=disable -silent' -WindowStyle Hidden

# Check if reboot will be needed and create task
Start-Sleep 5  # Give DCU time to assess updates
$rebootCheck = Start-Process -NoNewWindow -Wait -FilePath $DCU -ArgumentList '/rebootpending' -PassThru
if ($rebootCheck.ExitCode -eq 0) {
    Create-PostRebootTask
    Write-Output "Dell updates started - post-reboot task created"
} else {
    Write-Output "Dell updates started in background"
}
