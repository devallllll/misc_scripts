<#
.SYNOPSIS
    Disables computers inactive for a defined period (default 365 days) and moves them to a quarantine OU (ArchiveCleanup),
    plus generates an audit report of all computers currently in ArchiveCleanup with their last-seen timestamps.

.DESCRIPTION
    - Auto-detects the domain DN and default Computers container.
    - Ensures the ArchiveCleanup OU exists (creates it if missing).
    - Identifies ENABLED computers whose lastLogonTimestamp is older than the cutoff (default 365 days).
    - Disables and moves those computers into the ArchiveCleanup OU.
    - Writes two CSVs:
        1) Action report: which computers were disabled/moved (or would be in WhatIf).
        2) Audit report: all computers in ArchiveCleanup with LastSeen (derived from lastLogonTimestamp) and other metadata.
    - Supports -WhatIfMode to preview changes safely.

.PARAMETER InactivityDays
    Number of days a computer must be inactive to be considered stale. Default: 365.

.PARAMETER TargetOUName
    Name of the quarantine OU. Default: ArchiveCleanup.

.PARAMETER ExcludeServers
    If specified, excludes computers where OperatingSystem contains "Server".

.PARAMETER WhatIfMode
    If specified, performs a dry run (no changes), but still generates the reports.

.EXAMPLE
    .\Disable-And-Quarantine-StaleComputers.ps1 -WhatIfMode
    Preview which computers would be disabled and moved; also outputs the ArchiveCleanup audit report.

.EXAMPLE
    .\Disable-And-Quarantine-StaleComputers.ps1 -InactivityDays 400 -ExcludeServers
    Operates live: disables + moves workstations inactive for ≥400 days, excludes Server OS, writes reports.

.OUTPUTS
    CSV files in C:\Temp\
      - StaleComputers_ActionReport_<timestamp>.csv
      - ArchiveCleanup_Audit_<timestamp>.csv

.NOTES
    Author: GoodChoice IT Ltd
    Version: 1.0
    Tested on: Windows Server 2019/2022 (RSAT ActiveDirectory)
#>

param(
    [int]$InactivityDays = 365,
    [string]$TargetOUName = "ArchiveCleanup",
    [switch]$ExcludeServers,
    [switch]$WhatIfMode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module ActiveDirectory -ErrorAction Stop

# --- Paths & Time ---
$now = Get-Date
$stamp = $now.ToString('yyyy-MM-dd_HHmm')
$reportDir = 'C:\Temp'
$null = New-Item -ItemType Directory -Path $reportDir -Force -ErrorAction SilentlyContinue
$actionReport = Join-Path $reportDir ("StaleComputers_ActionReport_{0}.csv" -f $stamp)
$auditReport  = Join-Path $reportDir ("ArchiveCleanup_Audit_{0}.csv"  -f $stamp)

# --- Domain & OU setup ---
$domain      = Get-ADDomain
$domainDN    = $domain.DistinguishedName
$targetOU    = "OU=$TargetOUName,$domainDN"

# Ensure target OU exists
$ouExists = Get-ADOrganizationalUnit -LDAPFilter "(ou=$TargetOUName)" -SearchBase $domainDN -SearchScope Subtree -ErrorAction SilentlyContinue
if (-not $ouExists) {
    New-ADOrganizationalUnit -Name $TargetOUName -Path $domainDN -ProtectedFromAccidentalDeletion $true | Out-Null
    Write-Host "Created OU: $targetOU"
}

# --- Build stale cutoff ---
$cutoff = $now.AddDays(-[math]::Abs($InactivityDays))  # ensure positive days

# --- Gather enabled computers and filter by inactivity ---
# lastLogonTimestamp is a large integer (FILETIME). Convert to DateTime for comparison.
# Filter servers optionally.
$filter = '(objectCategory=computer)'
$props  = @('Name','DistinguishedName','OperatingSystem','LastLogonTimestamp','whenCreated')

$enabledComputers = Get-ADComputer -SearchBase $domainDN -SearchScope Subtree -LDAPFilter $filter -Properties $props, 'Enabled' -ResultPageSize 2000 |
    Where-Object { $_.Enabled -eq $true } |
    ForEach-Object {
        $lastSeen = $null
        if ($_.LastLogonTimestamp) {
            try { $lastSeen = [DateTime]::FromFileTime($_.LastLogonTimestamp) } catch { $lastSeen = $null }
        }
        # Construct an object including computed LastSeen
        [pscustomobject]@{
            Name               = $_.Name
            DistinguishedName  = $_.DistinguishedName
            OperatingSystem    = $_.OperatingSystem
            Enabled            = $_.Enabled
            WhenCreated        = $_.whenCreated
            LastSeen           = $lastSeen
        }
    }

if ($ExcludeServers) {
    $enabledComputers = $enabledComputers | Where-Object { $_.OperatingSystem -notmatch 'Server' }
}

$stale = $enabledComputers | Where-Object {
    # Consider stale if LastSeen exists and is older than cutoff.
    # If LastSeen is null, you can choose to treat as stale by uncommenting the second condition.
    ($_.LastSeen -and $_.LastSeen -lt $cutoff)
    # -or (-not $_.LastSeen)     # <- optional: treat "never logged on" as stale too
}

# --- Disable + Move stale computers (or WhatIf) ---
$actions = New-Object System.Collections.Generic.List[object]

foreach ($c in $stale) {
    $action = 'Moved+Disabled'
    $err    = $null

    try {
        if ($WhatIfMode) {
            Disable-ADAccount -Identity $c.DistinguishedName -WhatIf
            Move-ADObject    -Identity $c.DistinguishedName -TargetPath $targetOU -WhatIf
            $action = 'WouldMove+WouldDisable'
        } else {
            Disable-ADAccount -Identity $c.DistinguishedName -Confirm:$false
            Move-ADObject    -Identity $c.DistinguishedName -TargetPath $targetOU -Confirm:$false
        }
    } catch {
        $action = 'Error'
        $err = $_.Exception.Message
    }

    $actions.Add([pscustomobject]@{
        Name               = $c.Name
        DistinguishedName  = $c.DistinguishedName
        OperatingSystem    = $c.OperatingSystem
        LastSeen           = $c.LastSeen
        Cutoff             = $cutoff
        Timestamp          = $now
        Action             = $action
        Error              = $err
    })
}

# Write action report (ev
