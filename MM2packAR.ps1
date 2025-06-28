<#
.SYNOPSIS
    Packs all immediate subfolders of a dropped folder into a single .ar archive for Midtown Madness 2.

.DESCRIPTION
    This script launches a borderless, always-on-top WinForms GUI that accepts drag-and-drop of a root folder. 
    Upon dropping, it compresses all its immediate subdirectories into one ZIP file, then renames the file
    extension from .zip to .ar for compatibility with Midtown Madness 2.

.PARAMETER None
    The script uses a drag-and-drop graphical interface instead of command-line parameters.

.EXAMPLE
    PS> .\mm2packar.ps1
    Launches the GUI. Drag the folder named 'vehicle' onto the form to generate 'vehicle.ar' in the same
    parent directory.

.NOTES
    - Requires PowerShell 7+ and .NET Core.
    - Utilizes System.IO.Compression.FileSystem for ZIP operations.
    - The form is borderless, dark-themed, draggable, and positioned at the bottom-right of the screen.
    - Includes a custom close button and TopMost behavior.

.AUTHOR
    Developed by chasseyblue

.VERSION
    1.0.0

.LINK
    https://github.com/chasseyblue/MM2packAR
#>

# Load WinForms & Drawing assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Enable dragging using C#
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace Win32Functions {
    public static class Win32 {
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
    }
}
'@ -Language CSharp


# Create the form
$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = 'None'
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.ForeColor = [System.Drawing.Color]::White
$form.Text = 'Midtown Madness 2 PackAR'
$form.Width = 300                   
$form.Height = 200                  
$form.StartPosition = 'Manual'      # So we can move it later
$form.TopMost = $true               # Keep window above others
$form.AllowDrop = $true             # Enable drag-and-drop

# Get the usable desktop area (so we don’t overlap the taskbar)
$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea

# Compute X,Y so the form hugs the bottom-right, 
# with e.g. a 10px margin from edges
$margin = 10
$x = $wa.Width - $form.Width - $margin
$y = $wa.Height - $form.Height - $margin

# Apply the location
$form.Location = New-Object System.Drawing.Point($x, $y)

# THEN we show it as before:
[System.Windows.Forms.Application]::EnableVisualStyles()

# Close button
$btnClose = New-Object System.Windows.Forms.button
$btnClose.Text = 'x'
$btnClose.Font = New-Object System.Drawing.Font('Segoe UI',14,[System.Drawing.FontStyle]::Bold)
$btnClose.Size = New-Object System.Drawing.Size(30,30)
$btnClose.Location = New-Object System.Drawing.Point($form.ClientSize.Width - $btnClose.Width - $margin)
$btnClose.FlatStyle = 'Flat'
$btnClose.ForeColor = $form.ForeColor
$btnClose.BackColor = $form.BackColor
$btnClose.FlatAppearance.BorderSize = 0

# OnClick, exit form and kill process
$btnClose.Add_Click({$form.Close()})

# Place on top (close button)
$form.Controls.Add($btnClose)
$btnClose.BringToFront()

# Add a label for instructions
$label = New-Object System.Windows.Forms.Label
$label.BackColor = $form.BackColor
$label.ForeColor = $form.ForeColor
$label.Text = "Drag your vehicle's top folder here to create .AR"
$label.Dock = 'Fill'
$label.TextAlign = 'MiddleCenter'
$label.Font = New-Object System.Drawing.Font('Segoe UI', 12)
$form.Controls.Add($label)

# When you drag over, show Copy cursor if it's folders
$form.Add_DragEnter({
        if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
            $_.Effect = 'Copy'
        }
        else {
            $_.Effect = 'None'
        }
    })

# On drop: process each dropped path
$form.Add_DragDrop({
        $paths = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
        foreach ($path in $paths) {
            if (-not (Test-Path $path -PathType Container)) {
                [Windows.Forms.MessageBox]::Show(
                    "'$path' is not a folder.",
                    'Invalid Drop',
                    [Windows.Forms.MessageBoxButtons]::OK,
                    [Windows.Forms.MessageBoxIcon]::Warning
                )
                continue
            }

            $folder = Get-Item $path
            $zipPath = Join-Path $folder.Parent.FullName ("$($folder.BaseName).zip")
            $arPath = [IO.Path]::ChangeExtension($zipPath, '.ar')

            # ZIP and Compression logic
            try {
                # load the ZipFile & ZipArchive types
                Add-Type -AssemblyName System.IO.Compression.FileSystem

                # open a single archive for all subfolders
                $zip = [IO.Compression.ZipFile]::Open(
                    $zipPath,
                    [IO.Compression.ZipArchiveMode]::Create
                )

                # for each immediate subfolder...
                Get-ChildItem $folder.FullName -Directory | ForEach-Object {
                    $sub = $_
                    $baseName = $sub.Name

                    # enumerate every file inside that subfolder
                    Get-ChildItem $sub.FullName -Recurse -File | ForEach-Object {
                        $file = $_
                        # build a path inside the archive like "SubfolderName/…"
                        $relative = $file.FullName.Substring($sub.FullName.Length + 1)
                        $entryName = "$baseName/$relative" -replace '\\', '/'

                        # create the entry and copy file data in
                        $entryStream = $zip.CreateEntry($entryName).Open()
                        [IO.File]::OpenRead($file.FullName).CopyTo($entryStream)
                        $entryStream.Dispose()
                    }
                }

                $zip.Dispose()

                # rename .zip > .ar
                Rename-Item -Path $zipPath -NewName $arPath -Force

                [Windows.Forms.MessageBox]::Show(
                    "Packed subfolders into:`n$($arPath)",
                    'Success',
                    [Windows.Forms.MessageBoxButtons]::OK,
                    [Windows.Forms.MessageBoxIcon]::Information
                )
            }
            catch {
                [Windows.Forms.MessageBox]::Show(
                    "Error:`n$($_.Exception.Message)",
                    'Oops!',
                    [Windows.Forms.MessageBoxButtons]::OK,
                    [Windows.Forms.MessageBoxIcon]::Error
                )
            }
        }
    })

# Enable hook for dragging
    foreach ($ctrl in @($form, $label)) {
    $ctrl.Add_MouseDown({
            param($s, $e)
            if ($e.Button -eq [Windows.Forms.MouseButtons]::Left) {
                [Win32Functions.Win32]::ReleaseCapture()
                # 0xA1 = WM_NCLBUTTONDOWN, 0x2 = HTCAPTION
                [Win32Functions.Win32]::SendMessage($form.Handle, 0xA1, 0x2, 0)
            }
        })
}

# Show the form
[void]$form.ShowDialog()
