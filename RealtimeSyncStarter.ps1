Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------
# FORM
# ---------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "RealtimeSync Startup Manager"
$form.Size = New-Object System.Drawing.Size(520,420)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'

# ---- Dark mode ----
$darkBack  = [System.Drawing.Color]::FromArgb(32,32,32)
$darkGrid  = [System.Drawing.Color]::FromArgb(45,45,45)
$darkText  = [System.Drawing.Color]::White
$buttonBack = [System.Drawing.Color]::FromArgb(60,60,60)

$form.BackColor = $darkBack
$form.ForeColor = $darkText

# ---------------------------
# Startup folder
# ---------------------------

$startupFolder = [Environment]::GetFolderPath("Startup")

# ---------------------------
# GRID
# ---------------------------

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Top'
$grid.Height = 320
$grid.AllowUserToAddRows = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.RowHeadersVisible = $false
$grid.BackgroundColor = $darkGrid
$grid.DefaultCellStyle.BackColor = $darkGrid
$grid.DefaultCellStyle.ForeColor = $darkText
$grid.ColumnHeadersDefaultCellStyle.BackColor = $darkBack
$grid.ColumnHeadersDefaultCellStyle.ForeColor = $darkText
$grid.EnableHeadersVisualStyles = $false
$form.Controls.Add($grid)

# File column
$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = "FileName"
$colName.HeaderText = "File Name"
$colName.AutoSizeMode = "Fill"
$grid.Columns.Add($colName)

# Running column (small)
$colRunning = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colRunning.Name = "Running"
$colRunning.HeaderText = "On"
$colRunning.Width = 45
$colRunning.ReadOnly = $true
$grid.Columns.Add($colRunning)

# ---------------------------
# BUTTON PANEL
# ---------------------------

$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = 'Bottom'
$panel.Height = 60
$panel.BackColor = $darkBack
$form.Controls.Add($panel)

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Add"
$btnAdd.Width = 200
$btnAdd.Left = 30
$btnAdd.Top = 15
$btnAdd.BackColor = $buttonBack
$btnAdd.ForeColor = $darkText
$panel.Controls.Add($btnAdd)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = "Delete"
$btnDelete.Width = 200
$btnDelete.Left = 260
$btnDelete.Top = 15
$btnDelete.BackColor = $buttonBack
$btnDelete.ForeColor = $darkText
$panel.Controls.Add($btnDelete)

# ---------------------------
# RealtimeSync path
# ---------------------------

function Get-RealtimeSyncPath {
    $paths = @(
        "C:\Program Files\FreeFileSync\RealtimeSync.exe",
        "C:\Program Files (x86)\FreeFileSync\RealtimeSync.exe"
    )
    foreach ($p in $paths) {
        if (Test-Path $p) { return $p }
    }
    return $null
}

$realtimeSyncPath = Get-RealtimeSyncPath
if (!$realtimeSyncPath) {
    [System.Windows.Forms.MessageBox]::Show("RealtimeSync.exe not found.")
    exit
}

$WScriptShell = New-Object -ComObject WScript.Shell

# ---------------------------
# Process Detection
# ---------------------------

function Get-RunningProcessForBatch($batchPath) {
    $procs = Get-CimInstance Win32_Process -Filter "Name = 'RealtimeSync.exe'"
    foreach ($p in $procs) {
        if ($p.CommandLine -match [Regex]::Escape($batchPath)) {
            return $p
        }
    }
    return $null
}

function Start-JobIfNotRunning($batchPath) {
    $proc = Get-RunningProcessForBatch $batchPath
    if (-not $proc) {
        Start-Process $realtimeSyncPath -ArgumentList "`"$batchPath`""
        return $true
    }
    return $true
}

function Stop-Job($batchPath) {
    $proc = Get-RunningProcessForBatch $batchPath
    if ($proc) {
        Stop-Process -Id $proc.ProcessId -Force
    }
}

# ---------------------------
# Refresh Grid
# ---------------------------

function RefreshGrid {
    $grid.Rows.Clear()

    Get-ChildItem $startupFolder -Filter "*-RealtimeSync.lnk" | ForEach-Object {

        $shortcut = $WScriptShell.CreateShortcut($_.FullName)
        $batchFile = $shortcut.Arguments.Trim('"')

        $name = [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -replace '-RealtimeSync$',''
        $row = $grid.Rows.Add($name, $false)

        $proc = Get-RunningProcessForBatch $batchFile

        if ($proc) {
            $grid.Rows[$row].Cells["Running"].Value = $true
        }
        else {
            # Auto restart if crashed
            Start-JobIfNotRunning $batchFile
            $grid.Rows[$row].Cells["Running"].Value = $true
        }
    }
}

# ---------------------------
# Add Button
# ---------------------------

$btnAdd.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "FreeFileSync Batch|*.ffs_batch"
    $ofd.Multiselect = $true

    if ($ofd.ShowDialog() -eq "OK") {

        foreach ($file in $ofd.FileNames) {

            $name = [System.IO.Path]::GetFileNameWithoutExtension($file)
            $shortcutPath = Join-Path $startupFolder "$name-RealtimeSync.lnk"

            if (!(Test-Path $shortcutPath)) {
                $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $realtimeSyncPath
                $shortcut.Arguments = "`"$file`""
                $shortcut.WorkingDirectory = Split-Path $realtimeSyncPath
                $shortcut.Save()
            }

            Start-JobIfNotRunning $file
        }

        RefreshGrid
    }
})

# ---------------------------
# Delete Button
# ---------------------------

$btnDelete.Add_Click({

    if ($grid.SelectedRows.Count -eq 0) { return }

    foreach ($row in $grid.SelectedRows) {

        $fileName = $row.Cells["FileName"].Value
        $shortcutPath = Join-Path $startupFolder "$fileName-RealtimeSync.lnk"

        if (Test-Path $shortcutPath) {

            $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
            $batchFile = $shortcut.Arguments.Trim('"')

            Stop-Job $batchFile
            Remove-Item $shortcutPath -Force
        }
    }

    RefreshGrid
})

# ---------------------------
# Auto Refresh Timer (detect crash)
# ---------------------------

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 5000
$timer.Add_Tick({ RefreshGrid })
$timer.Start()

RefreshGrid
$form.ShowDialog()
