Register-ScheduledTask -TaskName "CleanPublicDesktopShortcuts" `
  -Trigger (New-ScheduledTaskTrigger -Daily -At 10:00AM) `
  -Action (New-ScheduledTaskAction -Execute "powershell.exe" -Argument '-ExecutionPolicy Bypass -Command "Remove-Item ''C:\Users\Public\Desktop\*.lnk'' -Force -ErrorAction SilentlyContinue"') `
  -Principal (New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest)
