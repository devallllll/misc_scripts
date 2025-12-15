# v1.0
# Extracts GPP Drive Maps from all GPOs + captures GPO links and security filtering.
# Output: CSV

Import-Module GroupPolicy -ErrorAction Stop
Import-Module ActiveDirectory -ErrorAction Stop

$OutDir = Join-Path $env:TEMP "GPO_DriveMap_Export"
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

$Results = New-Object System.Collections.Generic.List[object]

function Get-GpoSecurityFiltering {
    param([Microsoft.GroupPolicy.Gpo]$Gpo)

    # GPO security filtering: principals with GpoApply (and typically GpoRead)
    $perms = Get-GPPermission -Guid $Gpo.Id -All -ErrorAction SilentlyContinue
    if (-not $perms) { return @() }

    $apply = $perms | Where-Object {
        $_.Permission -eq "GpoApply"
    }

    return $apply | Select-Object -ExpandProperty Trustee | ForEach-Object {
        # Trustee sometimes returns "DOMAIN\Name"
        $_
    }
}

function Get-GpoLinks {
    param([Microsoft.GroupPolicy.Gpo]$Gpo)

    # This uses the XML report because it includes link targets
    $xmlPath = Join-Path $OutDir ("{0}.xml" -f $Gpo.Id)
    Get-GPOReport -Guid $Gpo.Id -ReportType Xml -Path $xmlPath

    [xml]$x = Get-Content $xmlPath -Raw

    $links = @()
    $linkNodes = $x.GPO.LinksTo
    if ($linkNodes -and $linkNodes.SOMPath) {
        # Sometimes it’s a single item, sometimes multiple
        foreach ($som in @($x.GPO.LinksTo)) {
            $links += [pscustomobject]@{
                SOMPath    = $som.SOMPath
                Enabled    = $som.Enabled
                NoOverride = $som.NoOverride
            }
        }
    }

    return ,$links
}

function Parse-DriveMapsFromGpoXml {
    param(
        [Microsoft.GroupPolicy.Gpo]$Gpo,
        [string]$XmlPath,
        [object[]]$Links,
        [string[]]$SecurityFiltering
    )

    [xml]$x = Get-Content $XmlPath -Raw

    # GPP Drive Maps live under:
    # User -> Preferences -> Drives
    # Computer -> Preferences -> Drives
    # In the XML report they appear as extension data; we search broadly for "Drives" preference items.
    $driveNodes = $x.SelectNodes("//*[local-name()='Drive' or local-name()='DriveMap' or local-name()='DriveMaps']")

    if (-not $driveNodes -or $driveNodes.Count -eq 0) {
        return
    }

    foreach ($dn in $driveNodes) {
        # Attempt to extract common attributes used by GPP drives
        $letter = $dn.letter
        if (-not $letter) { $letter = $dn.GetAttribute("letter") }

        $path = $dn.path
        if (-not $path) { $path = $dn.GetAttribute("path") }

        $action = $dn.action
        if (-not $action) { $action = $dn.GetAttribute("action") }

        $label = $dn.label
        if (-not $label) { $label = $dn.GetAttribute("label") }

        # Item-level targeting often contains group membership conditions; we’ll capture the raw ILT text if present.
        $ilt = $null
        $iltNode = $dn.SelectSingleNode(".//*[contains(translate(local-name(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'filter') or contains(translate(local-name(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'target')]")
        if ($iltNode) { $ilt = $iltNode.OuterXml }

        $scope = "Unknown"
        # crude scope detection from parent nodes in report
        $ancestor = $dn.SelectSingleNode("ancestor::*[contains(translate(local-name(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'user') or contains(translate(local-name(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'computer')]")
        if ($ancestor) {
            $n = $ancestor.LocalName.ToLowerInvariant()
            if ($n -like "*user*") { $scope = "User" }
            elseif ($n -like "*computer*") { $scope = "Computer" }
        }

        $Results.Add([pscustomobject]@{
            GPOName            = $Gpo.DisplayName
            GPOGuid            = $Gpo.Id.Guid
            Scope              = $scope
            DriveLetter        = $letter
            UNCPath            = $path
            Action             = $action
            Label              = $label
            ItemLevelTargeting = $ilt
            LinkedTo           = ($Links | ForEach-Object { $_.SOMPath }) -join " | "
            LinkEnabled        = ($Links | ForEach-Object { "$($_.SOMPath)=$($_.Enabled)" }) -join " | "
            SecurityFiltering  = ($SecurityFiltering -join " | ")
        })
    }
}

$AllGpos = Get-GPO -All
foreach ($g in $AllGpos) {
    try {
        $xmlPath = Join-Path $OutDir ("{0}.xml" -f $g.Id)
        $links = Get-GpoLinks -Gpo $g
        $sec = Get-GpoSecurityFiltering -Gpo $g

        Parse-DriveMapsFromGpoXml -Gpo $g -XmlPath $xmlPath -Links $links -SecurityFiltering $sec
    }
    catch {
        # Keep going; you can review failures later
        $Results.Add([pscustomobject]@{
            GPOName = $g.DisplayName
            GPOGuid = $g.Id.Guid
            Error   = $_.Exception.Message
        })
    }
}

$CsvPath = Join-Path $OutDir "GPO_DriveMappings.csv"
$Results | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $CsvPath

Write-Host "Done. Output:" $CsvPath
