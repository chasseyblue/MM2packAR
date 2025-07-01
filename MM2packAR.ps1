<#
.SYNOPSIS
    Packs all immediate subfolders of a dropped folder into a single .ar archive for Midtown Madness 2.

.DESCRIPTION
    This script launches a borderless, always-on-top WinForms GUI that accepts drag-and-drop of a root folder. 
    Upon dropping, it compresses all its immediate subdirectories into one ZIP file, then renames the file
    extension from .zip to .ar for compatibility with Midtown Madness 2.

.EXAMPLE
    PS> .\mm2packar.ps1
    Launches the GUI. Drag the folder named 'vehicle' onto the form to generate 'vehicle.ar'.

.NOTES
    - Requires PowerShell 7+ and .NET Core.
    - Uses System.IO.Compression + FileSystem for ZIP.
    - Borderless, dark theme, draggable, bottom-right anchored.

.AUTHOR
    Developed by chasseyblue

.VERSION
    1.0.1
#>

# Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Ensure ZIP types & enums exist when compiled to EXE
[Reflection.Assembly]::LoadWithPartialName("System.IO.Compression")        | Out-Null
[Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null

# P/Invoke for draggable form
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

# Build Form
$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = 'None'
$form.BackColor       = [System.Drawing.Color]::FromArgb(30,30,30)
$form.ForeColor       = [System.Drawing.Color]::White
$form.Text            = 'Midtown Madness 2 PackAR'
$form.Width           = 300
$form.Height          = 200
$form.StartPosition   = 'Manual'    # we'll position it ourselves
$form.TopMost         = $true        # always on top
$form.AllowDrop       = $true        # enable drag & drop

# Position bottom-right
[int] $margin  = 10
$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
[int] $posX    = $wa.Width  - [int]$form.Width  - $margin
[int] $posY    = $wa.Height - [int]$form.Height - $margin
$form.Location = New-Object System.Drawing.Point($posX, $posY)

# Enable WinForms visual styles
[System.Windows.Forms.Application]::EnableVisualStyles()

# Close Button
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text  = '×'
$btnClose.Font  = New-Object System.Drawing.Font('Segoe UI',14,[System.Drawing.FontStyle]::Bold)
$btnClose.Size  = New-Object System.Drawing.Size(30,30)
[int] $btnX     = [int]$form.ClientSize.Width - [int]$btnClose.Width - $margin
[int] $btnY     = $margin
$btnClose.Location = New-Object System.Drawing.Point($btnX, $btnY)
$btnClose.FlatStyle = 'Flat'
$btnClose.ForeColor = $form.ForeColor
$btnClose.BackColor = $form.BackColor
$btnClose.FlatAppearance.BorderSize = 0
$btnClose.Cursor     = 'Hand'
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)
$btnClose.BringToFront()

# Instruction Label
$label = New-Object System.Windows.Forms.Label
$label.BackColor   = $form.BackColor
$label.ForeColor   = $form.ForeColor
$label.Text        = "Drag your vehicle's top folder here to create .AR"
$label.Dock        = 'Fill'
$label.TextAlign   = 'MiddleCenter'
$label.Font        = New-Object System.Drawing.Font('Segoe UI',12)
$form.Controls.Add($label)

# DragEnter (use the enum, not a boolean or string)
$form.Add_DragEnter({
    param($s,$e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    } else {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})

# DragDrop (one .ar from all subfolders)
$form.Add_DragDrop({
    param($s,$e)
    $paths = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    foreach ($path in $paths) {
        if (-not (Test-Path $path -PathType Container)) {
            [System.Windows.Forms.MessageBox]::Show(
                "'$path' is not a folder.",
                'Invalid Drop',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            continue
        }

        $folder  = Get-Item $path
        $zipPath = Join-Path $folder.Parent.FullName "$($folder.BaseName).zip"
        $arPath  = [IO.Path]::ChangeExtension($zipPath, '.ar')

        try {
            # Open a single ZIP for all subfolders
            $zip = [System.IO.Compression.ZipFile]::Open(
                $zipPath,
                [System.IO.Compression.ZipArchiveMode]::Create
            )

            # Add each immediate subfolder
            Get-ChildItem $folder.FullName -Directory | ForEach-Object {
                $sub  = $_
                $base = $sub.Name
                Get-ChildItem $sub.FullName -Recurse -File | ForEach-Object {
                    $file      = $_
                    $rel       = $file.FullName.Substring($sub.FullName.Length + 1)
                    $entryName = "$base/$rel" -replace '\\','/'

                    # *** Only one argument: the entry name ***
                    $entry      = $zip.CreateEntry($entryName)
                    $entryStream = $entry.Open()
                    [IO.File]::OpenRead($file.FullName).CopyTo($entryStream)
                    $entryStream.Dispose()
                }
            }

            $zip.Dispose()
            Rename-Item -Path $zipPath -NewName $arPath -Force

            [System.Windows.Forms.MessageBox]::Show(
                "Packed subfolders into:`n$($arPath)",
                'Success',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Error:`n$($_.Exception.Message)",
                'Oops!',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})

# Make form & label draggable everywhere
foreach ($ctrl in @($form, $label)) {
    $ctrl.Add_MouseDown({
        param($s,$e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            [Win32Functions.Win32]::ReleaseCapture()
            [Win32Functions.Win32]::SendMessage($form.Handle,0xA1,0x2,0)
        }
    })
}

# Show it!
[void] $form.ShowDialog()
