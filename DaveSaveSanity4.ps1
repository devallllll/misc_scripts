# Copyright David Lane 2024 all rights reserved

# Function to find the highest version number of Office installed
function Get-LatestOfficeVersion {
    $officePath = "HKCU:\Software\Microsoft\Office"
    try {
        $versions = Get-ChildItem -Path $officePath -ErrorAction Stop | Where-Object { $_.PSChildName -match '^\d+\.\d+$' } | Select-Object -ExpandProperty PSChildName
        $highestVersion = $versions | Sort-Object { [Version]$_ } -Descending | Select-Object -First 1
        return $highestVersion
    } catch {
        Write-Host "No Microsoft Office installation found."
        return $null
    }
}

function Get-IncrementedPath {
    param (
        [string]$currentPath,
        [string]$appName
    )
    $yearPattern = '\b(20\d{2})\b'
    if ($currentPath -match $yearPattern) {
        $year = [int]$Matches[1]
        $newYear = $year + 1
        $basePath = $currentPath -replace $yearPattern, $newYear

        # Ensure only one instance of the application name is appended
        $appNamePattern = [regex]::Escape("\$appName") + '$'
        if ($basePath -notmatch $appNamePattern) {
            $newPath = Join-Path $basePath $appName
        } else {
            $newPath = $basePath
        }

        return $newPath
    } else {
        Write-Host "No year found in the current path: $currentPath"
        return $null
    }
}



function Get-OfficePathdavelane {
    param (
        [string]$officeVersion,
        [string]$appName,
        [string]$keyName
    )
    try {
        $path = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$officeVersion\$appName\Options" -Name $keyName -ErrorAction Stop
        return $path.$keyName
    } catch {
        Write-Host "Unable to find the current $keyName for $appName. It might not be set yet."
        return $null
    }
}

# Function to create the directory if it doesn't exist
function Create-DirectoryIfNeeded {
    param (
        [string]$path
    )
    if (-not (Test-Path -Path $path)) {
        New-Item -ItemType Directory -Path $path | Out-Null
        Write-Host "Created directory: $path"
    } else {
        Write-Host "Directory already exists: $path"
    }
}

# Main script
$latestOfficeVersion = Get-LatestOfficeVersion
if (-not $latestOfficeVersion) {
    exit
}

Write-Host "Latest Office version detected: ${latestOfficeVersion}"

$applications = @(
    @{ Name = "Word"; Key = "DOC-PATH" },
    @{ Name = "Excel"; Key = "DefaultPath" }
)

foreach ($app in $applications) {
    $currentPath = Get-OfficePathdavelane -officeVersion $latestOfficeVersion -appName $app['Name'] -keyName $app['Key']
    if ($currentPath) {
        Write-Host "$($app['Name'])'s current $($app['Key']) for Office ${latestOfficeVersion}: $currentPath"
        $newPath = Get-IncrementedPath -currentPath $currentPath -appName $app['Name']
        if ($newPath) {
            Write-Host "Proposed new path for $($app['Name']): $newPath"
            $userConfirmation = Read-Host "Do you want to update the path and create the directory if needed? (Y/N)"
            if ($userConfirmation -eq 'Y') {
                Create-DirectoryIfNeeded -path $newPath
                # Update the registry with the new path
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$latestOfficeVersion\$($app['Name'])\Options" -Name $app['Key'] -Value $newPath -Type ExpandString
                Write-Host "$($app['Key']) for $($app['Name']) updated to $newPath"
            } else {
                Write-Host "No changes made to $($app['Name'])'s $($app['Key'])."
            }
        }
    } else {
        Write-Host "$($app['Name'])'s $($app['Key']) for Office ${latestOfficeVersion} is not currently set."
    }
}



Write-Host "Script completed."
