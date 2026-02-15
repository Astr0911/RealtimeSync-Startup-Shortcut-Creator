Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "RealtimeSync Startup Creator"
$form.Size = New-Object System.Drawing.Size(500,200)
$form.AllowDrop = $true
$form.TopMost = $true

$label = New-Object System.Windows.Forms.Label
$label.Text = "Drag one or multiple .ffs_batch files here"
$label.Dock = "Fill"
$label.TextAlign = "MiddleCenter"
$form.Controls.Add($label)

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

$form.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [Windows.Forms.DragDropEffects]::Copy
    }
})

$form.Add_DragDrop({
    $files = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    $realtimeSyncPath = Get-RealtimeSyncPath

    if (!$realtimeSyncPath) {
        [System.Windows.Forms.MessageBox]::Show("RealtimeSync not found.")
        return
    }

    $startupFolder = [Environment]::GetFolderPath("Startup")
    $WScriptShell = New-Object -ComObject WScript.Shell

    foreach ($file in $files) {
        if ($file.ToLower().EndsWith(".ffs_batch")) {

            $name = [System.IO.Path]::GetFileNameWithoutExtension($file)
            $shortcutPath = Join-Path $startupFolder "$name-RealtimeSync.lnk"

            if (Test-Path $shortcutPath) {
                Remove-Item $shortcutPath -Force
            }

            $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = $realtimeSyncPath
            $shortcut.Arguments = "`"$file`""
            $shortcut.WorkingDirectory = Split-Path $realtimeSyncPath
            $shortcut.Save()
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Startup shortcuts created successfully!")
})

$form.ShowDialog()
