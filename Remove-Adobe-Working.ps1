# Adobe Creative Cloud Complete Removal Script
# Run as Administrator for best results

param(
    [switch]$Force = $false
)

Write-Host "Adobe Creative Cloud Complete Removal Tool" -ForegroundColor Green
Write-Host ("=" * 50) -ForegroundColor Yellow

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if (-not $isAdmin) {
    Write-Host "This script should be run as Administrator for best results!" -ForegroundColor Yellow
    Write-Host "Right-click PowerShell and select 'Run as administrator'" -ForegroundColor Yellow
    if (-not $Force) {
        $continue = Read-Host "Continue anyway? (y/N)"
        if ($continue -ne 'y' -and $continue -ne 'Y') {
            exit
        }
    }
}

# Adobe processes to kill
$adobeProcesses = @(
    'Creative Cloud', 'Creative Cloud Helper', 'CCXProcess',
    'CCLibrary', 'AdobeIPCBroker', 'Adobe Desktop Service',
    'AdobeUpdateService', 'AdobeCleanUpUtility', 'acrotray',
    'Acrobat', 'AcroRd32', 'Photoshop', 'Illustrator',
    'InDesign', 'AfterFX', 'Premiere Pro', 'Bridge',
    'Lightroom', 'Dreamweaver', 'Animate', 'Audition',
    'CoreSync', 'AdobeCollabSync', 'AdobeNotificationClient'
)

# Adobe services to stop
$adobeServices = @(
    'AdobeUpdateService', 'Adobe Genuine Monitor Service', 
    'Adobe Genuine Software Integrity Service', 'AGSService',
    'AGMService', 'AdobeARMservice', 'Adobe LM Service'
)

# Step 1: Kill Adobe processes
Write-Host ""
Write-Host "Stopping Adobe processes..." -ForegroundColor Cyan
$killedCount = 0

foreach ($process in $adobeProcesses) {
    try {
        $runningProcesses = Get-Process -Name $process -ErrorAction SilentlyContinue
        if ($runningProcesses) {
            $runningProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
            Write-Host "  Killed: $process" -ForegroundColor Green
            $killedCount++
        }
    }
    catch {
        # Process not running, continue
    }
}

if ($killedCount -gt 0) {
    Write-Host "Successfully killed $killedCount Adobe processes" -ForegroundColor Green
    Start-Sleep -Seconds 3
} else {
    Write-Host "No Adobe processes were running" -ForegroundColor Blue
}

# Step 2: Stop Adobe services
Write-Host ""
Write-Host "Stopping Adobe services..." -ForegroundColor Cyan
$stoppedCount = 0

foreach ($service in $adobeServices) {
    try {
        $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -eq 'Running') {
            Stop-Service -Name $service -Force -ErrorAction SilentlyContinue
            Set-Service -Name $service -StartupType Disabled -ErrorAction SilentlyContinue
            Write-Host "  Stopped and disabled: $service" -ForegroundColor Green
            $stoppedCount++
        }
    }
    catch {
        # Service not found, continue
    }
}

if ($stoppedCount -gt 0) {
    Write-Host "Successfully stopped $stoppedCount Adobe services" -ForegroundColor Green
} else {
    Write-Host "No Adobe services were running" -ForegroundColor Blue
}

# Step 3: Run CC Cleaner
Write-Host ""
Write-Host "Running Adobe CC Cleaner Tool..." -ForegroundColor Cyan
$ccCleanerPath = "C:\AdobeCreativeCloudCleanerTool_Win\AdobeCreativeCloudCleanerTool.exe"

if (Test-Path $ccCleanerPath) {
    Write-Host "Launching CC Cleaner..." -ForegroundColor Yellow
    Write-Host "The CC Cleaner will open - follow these steps:" -ForegroundColor Yellow
    Write-Host "   1. Select 'Creative Cloud for desktop'" -ForegroundColor White
    Write-Host "   2. Choose 'Remove All'" -ForegroundColor White
    Write-Host "   3. Click 'Clean'" -ForegroundColor White
    Write-Host "   4. Wait for completion" -ForegroundColor White
    
    Start-Process -FilePath $ccCleanerPath -Wait
    Write-Host "CC Cleaner completed" -ForegroundColor Green
} else {
    Write-Host "CC Cleaner not found at: $ccCleanerPath" -ForegroundColor Red
    Write-Host "Please download it from Adobe's website" -ForegroundColor Yellow
}

# Step 4: Clean registry entries
Write-Host ""
Write-Host "Cleaning Adobe registry entries..." -ForegroundColor Cyan
$registryPaths = @(
    "HKLM:\SOFTWARE\Adobe",
    "HKLM:\SOFTWARE\WOW6432Node\Adobe"
)

$cleanedRegistry = 0
foreach ($regPath in $registryPaths) {
    try {
        if (Test-Path $regPath) {
            Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  Removed registry path: $regPath" -ForegroundColor Green
            $cleanedRegistry++
        }
    }
    catch {
        Write-Host "  Couldn't remove registry path: $regPath" -ForegroundColor Yellow
    }
}

# Clean uninstall entries
try {
    $uninstallPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $adobeUninstalls = Get-ChildItem -Path $uninstallPath -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Adobe*" }
    foreach ($item in $adobeUninstalls) {
        Remove-Item -Path $item.PSPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  Removed uninstall entry: $($item.PSChildName)" -ForegroundColor Green
        $cleanedRegistry++
    }
}
catch {
    Write-Host "  Error cleaning uninstall entries" -ForegroundColor Yellow
}

Write-Host "Cleaned $cleanedRegistry registry entries" -ForegroundColor Green

# Step 5: Clean file system
Write-Host ""
Write-Host "Cleaning Adobe files and folders..." -ForegroundColor Cyan
$adobePaths = @(
    "C:\Program Files\Adobe",
    "C:\Program Files (x86)\Adobe", 
    "C:\Program Files\Common Files\Adobe",
    "C:\Program Files (x86)\Common Files\Adobe",
    "C:\ProgramData\Adobe",
    "$env:LOCALAPPDATA\Adobe",
    "$env:APPDATA\Adobe",
    "$env:USERPROFILE\Creative Cloud Files"
)

$cleanedFiles = 0
foreach ($path in $adobePaths) {
    if (Test-Path $path) {
        try {
            Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  Removed: $path" -ForegroundColor Green
            $cleanedFiles++
        }
        catch {
            Write-Host "  Couldn't remove: $path (may be in use)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  Path not found: $path" -ForegroundColor Blue
    }
}

Write-Host "Cleaned $cleanedFiles file system locations" -ForegroundColor Green

# Step 6: Final cleanup
Write-Host ""
Write-Host "Performing final cleanup..." -ForegroundColor Cyan

# Clear temp files
$tempPaths = @(
    "C:\Windows\Temp\Adobe*",
    "$env:LOCALAPPDATA\Temp\Adobe*"
)

foreach ($tempPath in $tempPaths) {
    try {
        Remove-Item -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  Cleared temp files: $tempPath" -ForegroundColor Green
    }
    catch {
        # Continue if temp files can't be removed
    }
}

Write-Host "Final cleanup completed!" -ForegroundColor Green

# Summary
Write-Host ""
Write-Host ("=" * 50) -ForegroundColor Yellow
Write-Host "Adobe Creative Cloud removal completed!" -ForegroundColor Green
Write-Host "Summary:" -ForegroundColor White
Write-Host "   Adobe processes terminated" -ForegroundColor Green
Write-Host "   Adobe services stopped and disabled" -ForegroundColor Green
Write-Host "   CC Cleaner executed" -ForegroundColor Green
Write-Host "   Registry entries cleaned" -ForegroundColor Green
Write-Host "   File system cleaned" -ForegroundColor Green
Write-Host ""
Write-Host "Please restart your computer to complete the removal process." -ForegroundColor Yellow

Write-Host ""
Write-Host "Press Enter to exit..."
Read-Host
