# MM2PackAR

A lightweight PowerShell WinForms utility to package all immediate subfolders of a dropped folder into a single `.ar` archive, ready for use with **Midtown Madness 2**.

---

## 🔍 Synopsis

Drag & drop a folder onto the floating window, and **MM2PackAR** will:

1. Zip each subfolder under the dropped folder into one archive.  
2. Rename the resulting `.zip` to `.ar`.  
3. Place the `.ar` alongside the original folder.

Perfect for quickly bundling vehicle hierarchy into MM2’s custom archive format (.AR).

---

## ✨ Features

- **Always-On-Top**: Floats above all windows for quick access.  
- **Borderless Dark Theme**: Sleek, minimal UI that blends into your workflow.  
- **Bottom-Right Docking**: Auto-positions in your system tray area with a small margin.  
- **Drag-and-Drop**: Simply drag a folder—no browsing dialogs needed.  
- **Custom Close Button**: One-click exit with a “×” button.  
- **Form-Wide Drag**: Click & drag anywhere on the form to reposition it.  

---

## 📋 Prerequisites

- **Windows 10 / 11**  
- **PowerShell 7+** (tested on 7.x)  
- **.NET Core**  
- *(Optional)* Execution policy set to allow local scripts:
  ```powershell
  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass
  or
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
  ```

---

## 🚀 Installation

1. Clone or download this repository.  
2. Ensure `mm2packar.ps1` is unblocked:
   ```powershell
   Unblock-File .\mm2packar.ps1
   ```  
3. (Optional) Compile to an EXE with **PS2EXE**:
   ```powershell
   Install-Module ps2exe -Scope CurrentUser
   Invoke-PS2EXE .\mm2packar.ps1 .\mm2packar.exe -noConsole
   ```

---

## ⚙️ Usage

1. Launch the script (or EXE):
   ```powershell
   .\mm2packar.ps1
   ```
2. drag a folder (e.g. `vehicle`) onto the window.  
3. Watch for the success dialog—your `.ar` appears next to the original folder.  
4. Click “×” or press **Esc** to close.

---

## 🛠️ Example

```powershell
PS C:\Projects\MM2PackAR> .\mm2packar.ps1
# Window appears in bottom-right

# Drag “C:\Projects\Vehicles\SuperCar” onto it
# > Creates “C:\Projects\Vehicles\SuperCar.ar”
```

---

## 🔧 Customization

- **Window Position**  
  Adjust the `$margin` or target `Screen.WorkingArea` logic to dock elsewhere.  
- **Theme Colors**  
  Tweak the `BackColor` / `ForeColor` ARGB values for your own palette.  
- **Close-Button Style**  
  Customize the font, size, or add hover effects by modifying the `$btnClose` properties.

---

> 😊 Enjoy faster, hassle-free archiving for Midtown Madness 2—and happy racing!
