Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Ensure-Dir([string]$p) {
  if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null }
}

function New-ShortHash([string]$text) {
  $sha1 = [System.Security.Cryptography.SHA1]::Create()
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($text.ToLowerInvariant())
  $hash = $sha1.ComputeHash($bytes)
  return (([BitConverter]::ToString($hash) -replace "-", "").Substring(0,10))
}

function Find-Categories([string]$root) {
  if (-not (Test-Path $root)) { return @() }
  Get-ChildItem -Path $root -Directory -Force |
    Select-Object -ExpandProperty Name |
    Sort-Object
}

function Append-Log([System.Windows.Forms.TextBox]$tb, [string]$msg) {
  $stamp = (Get-Date).ToString("HH:mm:ss")
  $tb.AppendText("[$stamp] $msg`r`n")
}

function Flatten-Categories {
  param(
    [string]$SourceRoot,
    [string]$DestRoot,
    [string[]]$Categories,
    [bool]$CleanDest,
    [System.Windows.Forms.ProgressBar]$Progress,
    [System.Windows.Forms.TextBox]$LogBox
  )

  if (-not (Test-Path $SourceRoot)) { throw "SourceRoot not found: $SourceRoot" }
  if ([string]::IsNullOrWhiteSpace($DestRoot)) { throw "DestRoot is empty." }
  if (-not $Categories -or $Categories.Count -eq 0) { throw "No categories selected." }

  if ($CleanDest -and (Test-Path $DestRoot)) {
    Append-Log $LogBox "Cleaning destination: $DestRoot"
    Remove-Item -Path $DestRoot -Recurse -Force
  }

  Ensure-Dir $DestRoot

  $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $logDir = Join-Path $DestRoot "_logs\$stamp"
  Ensure-Dir $logDir
  $mapCsv = Join-Path $logDir "flatten-map.csv"

  $rows = New-Object System.Collections.Generic.List[object]
  $totalFoldersToCopy = 0

  # First pass: count how many INF folders we will copy (for progress)
  foreach ($cat in $Categories) {
    $catSrc = Join-Path $SourceRoot $cat
    if (-not (Test-Path $catSrc)) { continue }

    $infFiles = Get-ChildItem -Path $catSrc -Recurse -Filter *.inf -File -Force -ErrorAction SilentlyContinue
    $infFolders = $infFiles | Select-Object -ExpandProperty DirectoryName -Unique
    $totalFoldersToCopy += $infFolders.Count
  }

  if ($totalFoldersToCopy -eq 0) {
    Append-Log $LogBox "No INF folders found in selected categories."
    return
  }

  $Progress.Minimum = 0
  $Progress.Maximum = $totalFoldersToCopy
  $Progress.Value = 0

  Append-Log $LogBox "Source: $SourceRoot"
  Append-Log $LogBox "Destination: $DestRoot"
  Append-Log $LogBox "Categories: $($Categories -join ', ')"
  Append-Log $LogBox "INF folders to copy: $totalFoldersToCopy"
  Append-Log $LogBox "----"

  $copied = 0

  foreach ($cat in $Categories) {
    $catSrc = Join-Path $SourceRoot $cat
    if (-not (Test-Path $catSrc)) {
      Append-Log $LogBox "Skipping missing category folder: $catSrc"
      continue
    }

    $catDest = Join-Path $DestRoot $cat
    Ensure-Dir $catDest

    $infFiles = Get-ChildItem -Path $catSrc -Recurse -Filter *.inf -File -Force -ErrorAction SilentlyContinue
    $infFolders = $infFiles | Select-Object -ExpandProperty DirectoryName -Unique

    Append-Log $LogBox "[$cat] INF folders found: $($infFolders.Count)"

    foreach ($folder in $infFolders) {
      $hash = New-ShortHash $folder
      $destFolder = Join-Path $catDest $hash
      Ensure-Dir $destFolder

      # Copy entire INF folder contents to keep INF+SYS+CAT+DLL together
      Copy-Item -Path (Join-Path $folder "*") -Destination $destFolder -Recurse -Force

      $rows.Add([pscustomobject]@{
        Category       = $cat
        SourceFolder   = $folder
        DestFolder     = $destFolder
        HashFolderName = $hash
      }) | Out-Null

      $copied++
      if ($copied -le $Progress.Maximum) { $Progress.Value = $copied }
      [System.Windows.Forms.Application]::DoEvents()
    }
  }

  $rows | Export-Csv -Path $mapCsv -NoTypeInformation -Encoding UTF8

  Append-Log $LogBox "----"
  Append-Log $LogBox "Done. Copied INF folders: $copied"
  Append-Log $LogBox "Map/log written: $mapCsv"
}

# ---------------- GUI ----------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Flatten Dell Driver Pack (Shallow)"
$form.Size = New-Object System.Drawing.Size(900, 650)
$form.StartPosition = "CenterScreen"

$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = "Source root (contains network/chipset/storage folders):"
$lblSource.Location = New-Object System.Drawing.Point(10, 15)
$lblSource.Size = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(10, 38)
$txtSource.Size = New-Object System.Drawing.Size(760, 24)
$form.Controls.Add($txtSource)

$btnBrowseSource = New-Object System.Windows.Forms.Button
$btnBrowseSource.Text = "Browse..."
$btnBrowseSource.Location = New-Object System.Drawing.Point(780, 36)
$btnBrowseSource.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnBrowseSource)

$lblDest = New-Object System.Windows.Forms.Label
$lblDest.Text = "Destination root (output will be shallow):"
$lblDest.Location = New-Object System.Drawing.Point(10, 75)
$lblDest.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($lblDest)

$txtDest = New-Object System.Windows.Forms.TextBox
$txtDest.Location = New-Object System.Drawing.Point(10, 98)
$txtDest.Size = New-Object System.Drawing.Size(760, 24)
$form.Controls.Add($txtDest)

$btnBrowseDest = New-Object System.Windows.Forms.Button
$btnBrowseDest.Text = "Browse..."
$btnBrowseDest.Location = New-Object System.Drawing.Point(780, 96)
$btnBrowseDest.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnBrowseDest)

$chkClean = New-Object System.Windows.Forms.CheckBox
$chkClean.Text = "Clean destination first"
$chkClean.Location = New-Object System.Drawing.Point(10, 135)
$chkClean.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($chkClean)

$btnLoadCats = New-Object System.Windows.Forms.Button
$btnLoadCats.Text = "Load categories from source"
$btnLoadCats.Location = New-Object System.Drawing.Point(230, 130)
$btnLoadCats.Size = New-Object System.Drawing.Size(210, 28)
$form.Controls.Add($btnLoadCats)

$lblCats = New-Object System.Windows.Forms.Label
$lblCats.Text = "Select categories to flatten:"
$lblCats.Location = New-Object System.Drawing.Point(10, 170)
$lblCats.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($lblCats)

$clbCats = New-Object System.Windows.Forms.CheckedListBox
$clbCats.Location = New-Object System.Drawing.Point(10, 195)
$clbCats.Size = New-Object System.Drawing.Size(320, 200)
$clbCats.CheckOnClick = $true
$form.Controls.Add($clbCats)

$btnSelectCommon = New-Object System.Windows.Forms.Button
$btnSelectCommon.Text = "Select common (network/chipset/storage)"
$btnSelectCommon.Location = New-Object System.Drawing.Point(350, 195)
$btnSelectCommon.Size = New-Object System.Drawing.Size(260, 30)
$form.Controls.Add($btnSelectCommon)

$btnSelectNone = New-Object System.Windows.Forms.Button
$btnSelectNone.Text = "Select none"
$btnSelectNone.Location = New-Object System.Drawing.Point(350, 235)
$btnSelectNone.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnSelectNone)

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Select all"
$btnSelectAll.Location = New-Object System.Drawing.Point(490, 235)
$btnSelectAll.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnSelectAll)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Flatten"
$btnRun.Location = New-Object System.Drawing.Point(350, 285)
$btnRun.Size = New-Object System.Drawing.Size(260, 40)
$form.Controls.Add($btnRun)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(10, 410)
$progress.Size = New-Object System.Drawing.Size(860, 20)
$form.Controls.Add($progress)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Log:"
$lblLog.Location = New-Object System.Drawing.Point(10, 440)
$lblLog.Size = New-Object System.Drawing.Size(60, 20)
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 465)
$txtLog.Size = New-Object System.Drawing.Size(860, 140)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

$folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog

$btnBrowseSource.Add_Click({
  if ($folderDlg.ShowDialog() -eq "OK") {
    $txtSource.Text = $folderDlg.SelectedPath
  }
})

$btnBrowseDest.Add_Click({
  if ($folderDlg.ShowDialog() -eq "OK") {
    $txtDest.Text = $folderDlg.SelectedPath
  }
})

$btnLoadCats.Add_Click({
  $clbCats.Items.Clear()
  $src = $txtSource.Text.Trim()
  if (-not (Test-Path $src)) {
    Append-Log $txtLog "Source folder not found."
    return
  }
  $cats = Find-Categories $src
  foreach ($c in $cats) { [void]$clbCats.Items.Add($c, $false) }
  Append-Log $txtLog "Loaded categories: $($cats -join ', ')"
})

$btnSelectCommon.Add_Click({
  $common = @("network","chipset","storage")
  for ($i=0; $i -lt $clbCats.Items.Count; $i++) {
    $name = [string]$clbCats.Items[$i]
    $clbCats.SetItemChecked($i, ($common -contains $name.ToLowerInvariant()))
  }
})

$btnSelectNone.Add_Click({
  for ($i=0; $i -lt $clbCats.Items.Count; $i++) { $clbCats.SetItemChecked($i, $false) }
})

$btnSelectAll.Add_Click({
  for ($i=0; $i -lt $clbCats.Items.Count; $i++) { $clbCats.SetItemChecked($i, $true) }
})

$btnRun.Add_Click({
  try {
    $txtLog.Clear()

    $src = $txtSource.Text.Trim()
    $dst = $txtDest.Text.Trim()
    $clean = [bool]$chkClean.Checked

    $selected = @()
    foreach ($item in $clbCats.CheckedItems) { $selected += [string]$item }

    Append-Log $txtLog "Starting..."
    Flatten-Categories -SourceRoot $src -DestRoot $dst -Categories $selected -CleanDest $clean -Progress $progress -LogBox $txtLog
  }
  catch {
    Append-Log $txtLog "ERROR: $($_.Exception.Message)"
  }
})

[void]$form.ShowDialog()
