# --- PowerShell spawn monitor (no logging) ---

# Clean up any previous monitor with same name
Get-EventSubscriber -SourceIdentifier PSMonitor -ErrorAction SilentlyContinue |
    Unregister-Event -Force

Register-WmiEvent -Class Win32_ProcessStartTrace -SourceIdentifier PSMonitor -Action {
    $n = $Event.SourceEventArgs.NewEvent

    # Only care about PowerShell processes
    if ($n.ProcessName -in 'powershell.exe','pwsh.exe') {

        $procPid = $n.ProcessID

        $proc = Get-CimInstance Win32_Process -Filter "ProcessId = $procPid"
        $parent = Get-CimInstance Win32_Process -Filter "ProcessId = $($proc.ParentProcessId)"

        Write-Host ""
        Write-Host "=== NEW POWERSHELL SPAWNED ===" -ForegroundColor Yellow
        Write-Host (" Time:    {0}" -f (Get-Date))
        Write-Host (" PID:     {0}" -f $proc.ProcessId)
        Write-Host (" PS Cmd:  {0}" -f $proc.CommandLine)
        Write-Host (" Parent:  {0} (PID {1})" -f $parent.Name, $parent.ProcessId)
        Write-Host (" P Cmd:   {0}" -f $parent.CommandLine)
        Write-Host "====================================="
    }
}

Write-Host "Monitoring for new PowerShell processes..."
Write-Host "Leave this window open. Press Ctrl+C to stop." -ForegroundColor Cyan

while ($true) {
    Wait-Event -SourceIdentifier PSMonitor | Out-Null
}
