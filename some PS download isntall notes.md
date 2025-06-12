Notes on generic installer download

$url = "https://github.com/greenshot/greenshot/releases/download/v1.3.290/Greenshot-INSTALLER-1.3.290-RC1.exe"; $filename = Split-Path $url -Leaf; $file = "$env:TEMP\$filename"; Invoke-WebRequest -Uri $url -OutFile $file; Start-Process -FilePath $file -ArgumentList "/verysilent" -Wait
