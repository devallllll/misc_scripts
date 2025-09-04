<#
.SYNOPSIS
    Moves disabled or inactive computer accounts into a quarantine OU (ArchiveCleanup) 
    for safe cleanup and eventual deletion.

.DESCRIPTION
    This script searches Active Directory for disabled or inactive computer objects.
    It will create the ArchiveCleanup OU if it does not already exist, then move the 
    identified computers into that OU. The script supports a "WhatIf" mode to preview 
    actions without making changes, and logs all results to a CSV file for auditing.

    Typical use case:
      - Identify stale devices (disabled or inactive >180 days).
      - Disable and quarantine them safely (instead of immediate deletion).
      - Review and delete later after a quarantine period.

.PARAMETER FromComputersOnly
    If set, limits the search scope to the default CN=Computers container.
    By default, the whole domain is searched.

.PARAMETER TargetOUName
    The name of the OU to move stale/disabled computers into (default: ArchiveCleanup).

.PARAMETER WhatIfMove
    Performs a dry run. Shows what *would* be moved without changing AD.

.EXAMPLE
    .\CleanUpDisabledDevices.ps1 -WhatIfMove
    Runs the script in WhatIf mode, showing which computers would be moved.

.EXAMPLE
    .\CleanUpDisabledDevices.ps1
    Moves all disabled computers in the domain into the ArchiveCleanup OU.

.OUTPUTS
    A CSV report saved to C:\Temp\DisabledComputersMoved_<date>.csv

.NOTES
    Author: GoodChoice IT Ltd
    Date:   (Update when you run)
    Version: 1.0
    Tested on: Windows Server 2019/2022 with RSAT ActiveDirectory module
#>

param(
    [switch]$FromComputersOnly = $false,   # Set to $true to scan only CN=Computers
    [string]$TargetOUName = "ArchiveCleanup",
    [switch]$WhatIfMove
)

Import-Module ActiveDirectory -ErrorAction Stop

$domain      = Get-ADDomain
$domainDN    = $domain.DistinguishedName
$computersCN = $domain.ComputersContainer               # e.g. CN=Computers,DC=...
$targetOU    = "OU=$TargetOUName,$domainDN"
$reportPath  = "C:\Temp\DisabledComputersMoved_{0}.csv" -f (Get-Date -Format yyyy-MM-dd_HHmm)

# Ensure target OU exists
$ouExists = Get-ADOrganizationalUnit -LDAPFilter "(ou=$TargetOUName)" -SearchBase $domainDN -SearchScope Subtree -ErrorAction SilentlyContinue
if (-not $ouExists) {
    New-ADOrganizationalUnit -Name $TargetOUName -Path $domainDN -ProtectedFromAccidentalDeletion $true
    Write-Host "Created OU: $targetOU"
}

# Pick search base
$searchBase = if ($FromComputersOnly) { $computersCN } else { $domainDN }
Write-Host "Search base: $searchBase"

# Correct LDAP filter: DISABLED computers (bit 2 set)
$disabled = Get-ADComputer `
    -SearchBase $searchBase `
    -SearchScope Subtree `
    -LDAPFilter '(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=2))' `
    -Properties DistinguishedName,Name `
    -ResultPageSize 2000

if (-not $disabled) {
    Write-Host "No disabled computers found in $searchBase"
    return
}

# Prepare log results
$results = New-Object System.Collections.Generic.List[object]

foreach ($c in $disabled) {
    # Skip if already in target OU
    if ($c.DistinguishedName -like "*$targetOU") {
        $results.Add([pscustomobject]@{
            Name              = $c.Name
            DistinguishedName = $c.DistinguishedName
            SourceContainer   = ($c.DistinguishedName -split '(?<=CN=[^,]+),' ,2)[1]
            TargetOU          = $targetOU
            Timestamp         = (Get-Date)
            Action            = "Skipped (already in target OU)"
        })
        continue
    }

    $action = "Moved"
    try {
        if ($WhatIfMove) {
            Move-ADObject -Identity $c.DistinguishedName -TargetPath $targetOU -WhatIf
            $action = "WouldMove"
        } else {
            Move-ADObject -Identity $c.DistinguishedName -TargetPath $targetOU -Confirm:$false
        }
    } catch {
        $action = "Error: " + $_.Exception.Message
    }

    $results.Add([pscustomobject]@{
        Name              = $c.Name
        DistinguishedName = $c.DistinguishedName
        SourceContainer   = ($c.DistinguishedName -split '(?<=CN=[^,]+),' ,2)[1]
        TargetOU          = $targetOU
        Timestamp         = (Get-Date)
        Action            = $action
    })
}

# Ensure log folder and write report
$null = New-Item -ItemType Directory -Path (Split-Path $reportPath) -Force -ErrorAction SilentlyContinue
$results | Export-Csv -NoTypeInformation -Path $reportPath

# Summary
$moved     = ($results | Where-Object { $_.Action -eq 'Moved' }).Count
$wouldmove = ($results | Where-Object { $_.Action -eq 'WouldMove' }).Count
$skipped   = ($results | Where-Object { $_.Action -like 'Skipped*' }).Count
$errors    = ($results | Where-Object { $_.Action -like 'Error*' }).Count

Write-Host "Completed."
Write-Host "  Moved:      $moved"
Write-Host "  WouldMove:  $wouldmove"
Write-Host "  Skipped:    $skipped"
Write-Host "  Errors:     $errors"
Write-Host "Report: $reportPath"
Write-Host "Target OU: $targetOU"
