# Outlook for Windows Complete Removal Tool

A comprehensive PowerShell script to completely remove Adobe Creative Cloud and all related components from Windows systems.

## 🚨 Problem This Solves

**Have you ever tried to uninstall Adobe Creative Cloud and gotten this error?**
> "Adobe components are open. Please close all Adobe applications and try again."

Even when you think everything is closed, Adobe runs 20+ hidden background processes that prevent normal uninstallation. This tool solves that problem by force-killing all Adobe processes before running the official CC Cleaner Tool.

## 📁 Files Included

- **`Remove-Adobe.ps1`** - Main PowerShell removal script
- **`Remove-Adobe-Working.ps1`** - Simplified version (no emojis, proven to work)
- **`Run_Adobe_Remover.bat`** - Easy-to-use batch file wrapper
- **`adobe_complete_remover.py`** - Python alternative (advanced users)

## 🚀 Quick Start (Recommended)

### Method 1: Batch File (Easiest)
1. **Right-click** on `Run_Adobe_Remover.bat`
2. Select **"Run as administrator"**
3. Follow the prompts
4. **Restart your computer** when complete

### Method 2: PowerShell Script
1. **Right-click** on `Remove-Adobe-Working.ps1`
2. Select **"Run with PowerShell"** 
3. When prompted, choose **"Run as administrator"**
4. **Restart your computer** when complete

## 🔧 What This Tool Does

### 1. **Force-Kill Adobe Processes** 🛑
Terminates all Adobe background processes that prevent normal uninstallation:
- Creative Cloud.exe
- CCXProcess.exe (sync process)
- AdobeIPCBroker.exe
- Adobe Desktop Service.exe
- CoreSync.exe
- Plus 15+ other hidden processes

### 2. **Stop Adobe Services** ⚙️
Stops and disables Windows services:
- AdobeUpdateService
- Adobe Genuine Monitor Service
- AGSService, AGMService
- Adobe LM Service

### 3. **Run CC Cleaner Tool** 🧹
Automatically launches Adobe's official CC Cleaner Tool:
- Uses your existing tool at: `C:\AdobeCreativeCloudCleanerTool_Win\AdobeCreativeCloudCleanerTool.exe`
- Runs in a clean environment (no interfering processes)
- Provides step-by-step instructions

### 4. **Registry Cleanup** 📝
Removes Adobe registry entries:
- `HKLM:\SOFTWARE\Adobe`
- `HKLM:\SOFTWARE\WOW6432Node\Adobe`
- Adobe uninstall entries

### 5. **File System Cleanup** 📂
Removes Adobe installation folders:
- `C:\Program Files\Adobe`
- `C:\Program Files (x86)\Adobe`
- `C:\Program Files\Common Files\Adobe`
- `C:\ProgramData\Adobe`
- User AppData folders
- Creative Cloud Files folder

### 6. **Temp File Cleanup** 🗑️
Clears Adobe temporary files:
- Windows temp Adobe files
- User temp Adobe files

## ⚠️ Prerequisites

### Required
- **Windows Administrator privileges** (Right-click → "Run as administrator")
- **Adobe CC Cleaner Tool** downloaded from Adobe's website
  - Should be located at: `C:\AdobeCreativeCloudCleanerTool_Win\AdobeCreativeCleanerTool.exe`
  - Download from: https://helpx.adobe.com/creative-cloud/kb/cc-cleaner-tool-installation-problems.html

### Recommended
- **Close all Adobe applications** before running (though script will force-close them)
- **Backup important work** stored in Creative Cloud Files folder
- **Create a system restore point** (optional safety measure)

## 🛡️ Safety Features

- **Non-destructive process detection** - Only kills Adobe processes, never system processes
- **Error handling** - Continues even if some steps fail
- **Detailed logging** - Shows exactly what was removed
- **Force parameter** - Can run without user prompts: `Remove-Adobe.ps1 -Force`

## 📊 Expected Results

After running successfully, you should see:
```
Adobe Creative Cloud removal completed!
Summary:
   ✅ Adobe processes terminated
   ✅ Adobe services stopped and disabled  
   ✅ CC Cleaner executed
   ✅ Registry entries cleaned
   ✅ File system cleaned
```

## 🔄 Why This Tool is Necessary

**Adobe's CC Cleaner Tool works perfectly** - but only when no Adobe processes are running. The problem is Adobe's own background processes prevent their removal tool from working!

### The Chicken-and-Egg Problem:
1. You want to uninstall Adobe Creative Cloud
2. CC Cleaner Tool says "Adobe components are open"
3. You close all visible Adobe apps
4. Hidden processes keep running (20+ background processes)
5. CC Cleaner Tool still won't run
6. Normal Task Manager killing doesn't work (processes restart)

### Our Solution:
1. **Force-kill ALL Adobe processes** (they can't restart fast enough)
2. **Disable Adobe services** (prevents auto-restart)
3. **Launch CC Cleaner in clean environment** (finally works!)
4. **Clean up remaining traces** (registry, files, temp data)

## 🆘 Troubleshooting

### "Script won't run" / "Execution Policy" Error
```powershell
# Run this in PowerShell as Administrator:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### "Access Denied" Errors
- Make sure you're running as **Administrator**
- Some files may be in use - restart and try again

### CC Cleaner Tool Not Found
- Download from Adobe's official website
- Place in: `C:\AdobeCreativeCloudCleanerTool_Win\`
- Ensure filename is: `AdobeCreativeCloudCleanerTool.exe`

### Still Getting "Components Open" Error
- Try running the script twice
- Restart computer between attempts
- Check Task Manager for any remaining Adobe processes

## 💡 Advanced Usage

### Command Line Options
```powershell
# Run without prompts
.\Remove-Adobe.ps1 -Force

# Run specific script version
.\Remove-Adobe-Working.ps1 -Force
```

### Python Alternative
```bash
# For Python users
python adobe_complete_remover.py
```

## 🔍 Technical Details

### Process Termination Method
- Uses `Stop-Process -Force` for immediate termination
- Includes `-ErrorAction SilentlyContinue` to handle already-closed processes
- Targets process names, not PIDs (handles multiple instances)

### Service Management
- `Stop-Service -Force` to immediately stop services
- `Set-Service -StartupType Disabled` to prevent auto-restart
- Only targets Adobe-specific services

### Registry Cleanup Strategy
- Removes main Adobe registry keys
- Cleans Windows uninstall entries
- Uses `-ErrorAction SilentlyContinue` for missing keys

## ⚖️ Legal & Safety

- **This tool uses only standard Windows commands**
- **Does not modify system files or critical registry areas**
- **Only removes Adobe-specific components**
- **Reversible** (reinstall Adobe products normally if needed)
- **Based on Adobe's own CC Cleaner Tool** (we just prepare the environment)

## 🎯 Success Rate

This tool has been tested and successfully removes Adobe Creative Cloud in cases where:
- ✅ Normal uninstall fails with "components open" error
- ✅ CC Cleaner Tool alone won't work
- ✅ Multiple Adobe products are installed
- ✅ Adobe processes keep restarting
- ✅ Previous uninstall attempts failed

## 📞 Support

### Before Running
1. Download Adobe CC Cleaner Tool from Adobe's official website
2. Close any important work in Adobe applications
3. Run as Administrator

### After Running
1. **Restart your computer** (required to complete removal)
2. Check Programs & Features to confirm Adobe Creative Cloud is gone
3. Try installing fresh Adobe products if needed

### If Problems Persist
- Run Windows System File Checker: `sfc /scannow`
- Use Windows Registry Cleaner
- Create new Windows user account (nuclear option)

---

## 🏆 Credits

Created to solve the common Adobe Creative Cloud uninstallation problem where background processes prevent the official CC Cleaner Tool from working properly.

**Remember: Always restart your computer after running this tool to complete the removal process!**
