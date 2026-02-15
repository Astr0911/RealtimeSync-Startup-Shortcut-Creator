Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "RealtimeSync Startup Manager"
$form.Size = New-Object System.Drawing.Size(350,400)
$form.FormBorderStyle = 'FixedDialog'
$form.TopMost = $true
$form.StartPosition = 'CenterScreen'

$startupFolder = [Environment]::GetFolderPath("Startup")

# DataGridView for file name + running status
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Top'
$grid.Height = 300
$grid.AutoSizeColumnsMode = 'Fill'
$grid.AllowUserToAddRows = $false
$grid.SelectionMode = 'FullRowSelect'
$form.Controls.Add($grid)

# Add columns
$grid.Columns.Add("FileName","File Name") | Out-Null
$colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colCheck.Name = "Running"
$colCheck.HeaderText = "Running"
$colCheck.ReadOnly = $true
$grid.Columns.Add($colCheck) | Out-Null

# Buttons panel
$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = 'Bottom'
$panel.Height = 50
$form.Controls.Add($panel)

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Add"
$btnAdd.Width = 150
$btnAdd.Left = 20
$btnAdd.Top = 10
$panel.Controls.Add($btnAdd)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = "Delete"
$btnDelete.Width = 150
$btnDelete.Left = 180
$btnDelete.Top = 10
$panel.Controls.Add($btnDelete)

# Helper: find RealtimeSync.exe
function Get-RealtimeSyncPath {
    $paths = @(
        "C:\Program Files\FreeFileSync\RealtimeSync.exe",
        "C:\Program Files (x86)\FreeFileSync\RealtimeSync.exe"
    )
    foreach ($p in $paths) { if (Test-Path $p) { return $p } }
    return $null
}

$realtimeSyncPath = Get-RealtimeSyncPath
if (!$realtimeSyncPath) { [System.Windows.Forms.MessageBox]::Show("RealtimeSync not found."); exit }

$WScriptShell = New-Object -ComObject WScript.Shell

# Track running processes
$processes = @{}

# Refresh the grid with existing shortcuts
function RefreshGrid {
    $grid.Rows.Clear()
    Get-ChildItem $startupFolder -Filter "*-RealtimeSync.lnk" | ForEach-Object {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName) -replace '-RealtimeSync$',''
        $rowIndex = $grid.Rows.Add($name, $false)
        
        # Auto-start each job if not already running
        if (-not $processes.ContainsKey($_.FullName)) {
            $proc = Start-Process $_.FullName -PassThru
            $processes[$_.FullName] = $proc
            $grid.Rows[$rowIndex].Cells["Running"].Value = $true
        }
    }
}

RefreshGrid

# Add button: open file dialog
$btnAdd.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "FreeFileSync Batch|*.ffs_batch"
    $ofd.Multiselect = $true
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($file in $ofd.FileNames) {
            $name = [System.IO.Path]::GetFileNameWithoutExtension($file)
            $shortcutPath = Join-Path $startupFolder "$name-RealtimeSync.lnk"
            
            if (Test-Path $shortcutPath) { 
                # Stop existing process if present
                if ($processes.ContainsKey($shortcutPath)) {
                    $processes[$shortcutPath].Kill()
                    $processes.Remove($shortcutPath)
                }
                Remove-Item $shortcutPath -Force 
            }

            $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = $realtimeSyncPath
            $shortcut.Arguments = "`"$file`""
            $shortcut.WorkingDirectory = Split-Path $realtimeSyncPath
            $shortcut.Save()
        }
        RefreshGrid
    }
})

# Delete button
$btnDelete.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) { return }
    foreach ($row in $grid.SelectedRows) {
        $fileName = $row.Cells["FileName"].Value
        $shortcutPath = Join-Path $startupFolder "$fileName-RealtimeSync.lnk"
        
        if (Test-Path $shortcutPath) {
            # Kill process if running
            if ($processes.ContainsKey($shortcutPath)) {
                $processes[$shortcutPath].Kill()
                $processes.Remove($shortcutPath)
            }
            Remove-Item $shortcutPath -Force
        }
    }
    RefreshGrid
})

$form.ShowDialog()
