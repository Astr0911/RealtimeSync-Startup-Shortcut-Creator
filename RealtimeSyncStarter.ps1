Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Runtime.InteropServices

# ---------------------------
# Enable Windows 10/11 Dark Title Bar
# ---------------------------
$DWMWA_USE_IMMERSIVE_DARK_MODE = 20

$signature = @"
using System;
using System.Runtime.InteropServices;
public class Dwm {
    [DllImport("dwmapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int DwmSetWindowAttribute(
        IntPtr hwnd,
        int attr,
        ref int attrValue,
        int attrSize);
}
"@

Add-Type $signature

# ---------------------------
# FORM
# ---------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "RealtimeSync Startup Manager"
$form.Size = New-Object System.Drawing.Size(500,420)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'

# Dark colors
$darkBack  = [System.Drawing.Color]::FromArgb(32,32,32)
$darkGrid  = [System.Drawing.Color]::FromArgb(45,45,45)
$darkText  = [System.Drawing.Color]::White
$buttonBack = [System.Drawing.Color]::FromArgb(60,60,60)

$form.BackColor = $darkBack
$form.ForeColor = $darkText

# Apply dark title bar when shown
$form.Add_Shown({
    $value = 1
    [Dwm]::DwmSetWindowAttribute(
        $form.Handle,
        $DWMWA_USE_IMMERSIVE_DARK_MODE,
        [ref]$value,
        4
    ) | Out-Null
})

# ---------------------------
# Startup folder
# ---------------------------

$startupFolder = [Environment]::GetFolderPath("Startup")

# ---------------------------
# GRID
# ---------------------------

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Top'
$grid.Height = 310
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
# BUTTON PANEL (Vertical)
# ---------------------------

$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = 'Bottom'
$panel.Height = 80
$panel.BackColor = $darkBack
$form.Controls.Add($panel)

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Add"
$btnAdd.Dock = 'Top'
$btnAdd.Height = 35
$btnAdd.BackColor = $buttonBack
$btnAdd.ForeColor = $darkText
$panel.Controls.Add($btnAdd)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = "Delete"
$btnDelete.Dock = 'Top'
$btnDelete.Height = 35
$btnDelete.BackColor = $buttonBack
$btnDelete.ForeColor = $darkText
$panel.Controls.Add($btnDelete)

# ---------------------------
# Tray Icon
# ---------------------------

$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Icon = [System.Drawing.SystemIcons]::Application
$trayIcon.Text = "RealtimeSync Manager"
$trayIcon.Visible = $false

$trayMenu = New-Object System.Windows.Forms.ContextMenuStrip

$menuOpen = $trayMenu.Items.Add("Open")
$menuQuit = $trayMenu.Items.Add("Quit")

$trayIcon.ContextMenuStrip = $trayMenu

$menuOpen.Add_Click({
    $form.Show()
    $form.WindowState = "Normal"
})

$menuQuit.Add_Click({
    $trayIcon.Visible = $false
    $form.Close()
})

# Minimize to tray instead of closing
$form.Add_Resize({
    if ($form.WindowState -eq "Minimized") {
        $form.Hide()
        $trayIcon.Visible = $true
    }
})

# Double-click tray to restore
$trayIcon.Add_DoubleClick({
    $form.Show()
    $form.WindowState = "Normal"
})

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
# Process Handling
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
    }
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
        $rowIndex = $grid.Rows.Add($name, $false)

        $proc = Get-RunningProcessForBatch $batchFile

        if ($proc) {
            $grid.Rows[$rowIndex].Cells["Running"].Value = $true
        }
        else {
            Start-JobIfNotRunning $batchFile
            $grid.Rows[$rowIndex].Cells["Running"].Value = $true
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
# Crash Detection Timer
# ---------------------------

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 5000
$timer.Add_Tick({ RefreshGrid })
$timer.Start()

RefreshGrid
$form.ShowDialog()
