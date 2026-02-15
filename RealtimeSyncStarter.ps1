Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "RealtimeSync Startup Manager"
$form.Size = New-Object System.Drawing.Size(600,400)
$form.TopMost = $true

$startupFolder = [Environment]::GetFolderPath("Startup")

# ListBox for existing shortcuts
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Dock = "Top"
$listBox.Height = 250
$form.Controls.Add($listBox)

# Buttons
$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = "Bottom"
$panel.Height = 50
$form.Controls.Add($panel)

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Add"
$btnAdd.Width = 100
$btnAdd.Left = 20
$btnAdd.Top = 10
$panel.Controls.Add($btnAdd)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run"
$btnRun.Width = 100
$btnRun.Left = 140
$btnRun.Top = 10
$panel.Controls.Add($btnRun)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = "Delete"
$btnDelete.Width = 100
$btnDelete.Left = 260
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

# Populate list of shortcuts
function RefreshList {
    $listBox.Items.Clear()
    Get-ChildItem $startupFolder -Filter "*-RealtimeSync.lnk" | ForEach-Object {
        $listBox.Items.Add($_.FullName)
    }
}
RefreshList

$WScriptShell = New-Object -ComObject WScript.Shell

# Add button functionality
$btnAdd.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "FreeFileSync Batch|*.ffs_batch"
    $ofd.Multiselect = $true
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($file in $ofd.FileNames) {
            $name = [System.IO.Path]::GetFileNameWithoutExtension($file)
            $shortcutPath = Join-Path $startupFolder "$name-RealtimeSync.lnk"
            if (Test-Path $shortcutPath) { Remove-Item $shortcutPath -Force }
            $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = $realtimeSyncPath
            $shortcut.Arguments = "`"$file`""
            $shortcut.WorkingDirectory = Split-Path $realtimeSyncPath
            $shortcut.Save()
        }
        RefreshList
    }
})

# Run button functionality
$btnRun.Add_Click({
    if ($listBox.SelectedItem) {
        Start-Process $listBox.SelectedItem
    }
})

# Delete button functionality
$btnDelete.Add_Click({
    if ($listBox.SelectedItem) {
        Remove-Item $listBox.SelectedItem -Force
        RefreshList
    }
})

$form.ShowDialog()
