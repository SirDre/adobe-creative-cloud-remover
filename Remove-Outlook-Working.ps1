# Outlook for Windows Complete Removal Script
# Run as Administrator for best results

param(
    [switch]$Force = $false
)

Write-Host "Outlook for Windows Complete Removal Tool" -ForegroundColor Green
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

# Outlook processes to kill
$outlookProcesses = @(
    'OUTLOOK', 'outlook', 'Outlook',
    'olk', 'OLMAPI32', 'MAPIPH',
    'SCANPST', 'SCANOST'
)

# Outlook services to stop
$outlookServices = @(
    'OutlookService', 'Microsoft Outlook'
)

# Step 1: Kill Outlook processes
Write-Host ""
Write-Host "Stopping Outlook processes..." -ForegroundColor Cyan
$killedCount = 0

foreach ($process in $outlookProcesses) {
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
    Write-Host "Successfully killed $killedCount Outlook processes" -ForegroundColor Green
    Start-Sleep -Seconds 3
} else {
    Write-Host "No Outlook processes were running" -ForegroundColor Blue
}

# Step 2: Stop Outlook services
Write-Host ""
Write-Host "Stopping Outlook services..." -ForegroundColor Cyan
$stoppedCount = 0

foreach ($service in $outlookServices) {
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
    Write-Host "Successfully stopped $stoppedCount Outlook services" -ForegroundColor Green
} else {
    Write-Host "No Outlook services were running" -ForegroundColor Blue
}

# Step 3: Uninstall Outlook
Write-Host ""
Write-Host "Uninstalling Outlook..." -ForegroundColor Cyan

# Get Office installation info
$officeApps = Get-WmiObject -Class Win32_Product -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Outlook*" -or $_.Name -like "*Office*" }

if ($officeApps) {
    Write-Host "Found Office/Outlook installations" -ForegroundColor Yellow
    foreach ($app in $officeApps) {
        if ($app.Name -like "*Outlook*") {
            Write-Host "  Uninstalling: $($app.Name)" -ForegroundColor White
            try {
                $app.Uninstall() | Out-Null
                Write-Host "  Uninstalled: $($app.Name)" -ForegroundColor Green
            }
            catch {
                Write-Host "  Could not uninstall: $($app.Name)" -ForegroundColor Yellow
            }
        }
    }
}

# Alternative: Try Click-to-Run uninstall
Write-Host "  Attempting Click-to-Run removal..." -ForegroundColor White
$clickToRunPath = "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe"
if (Test-Path $clickToRunPath) {
    try {
        $arguments = 'scenario=install scenariosubtype=ARP sourcetype=None productstoremove=OutlookRetail.16_en-us_x-none culture=en-us version.16=16.0'
        Start-Process -FilePath $clickToRunPath -ArgumentList $arguments -Wait -NoNewWindow -ErrorAction SilentlyContinue
        Write-Host "  Click-to-Run removal attempted" -ForegroundColor Green
    }
    catch {
        Write-Host "  Click-to-Run removal encountered issues" -ForegroundColor Yellow
    }
}

Write-Host "Uninstall process completed" -ForegroundColor Green

# Step 4: Clean registry entries
Write-Host ""
Write-Host "Cleaning Outlook registry entries..." -ForegroundColor Cyan
$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Office\Outlook",
    "HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook",
    "HKLM:\SOFTWARE\Microsoft\Office\15.0\Outlook",
    "HKCU:\SOFTWARE\Microsoft\Office\Outlook",
    "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook",
    "HKCU:\SOFTWARE\Microsoft\Office\15.0\Outlook"
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
    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    
    foreach ($uninstallPath in $uninstallPaths) {
        $outlookUninstalls = Get-ChildItem -Path $uninstallPath -ErrorAction SilentlyContinue | Where-Object { $_.GetValue("DisplayName") -like "*Outlook*" }
        foreach ($item in $outlookUninstalls) {
            Remove-Item -Path $item.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  Removed uninstall entry: $($item.GetValue('DisplayName'))" -ForegroundColor Green
            $cleanedRegistry++
        }
    }
}
catch {
    Write-Host "  Error cleaning uninstall entries" -ForegroundColor Yellow
}

Write-Host "Cleaned $cleanedRegistry registry entries" -ForegroundColor Green

# Step 5: Clean file system
Write-Host ""
Write-Host "Cleaning Outlook files and folders..." -ForegroundColor Cyan
$outlookPaths = @(
    "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
    "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
    "C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
    "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
    "$env:LOCALAPPDATA\Microsoft\Outlook",
    "$env:APPDATA\Microsoft\Outlook",
    "$env:LOCALAPPDATA\Microsoft\Windows\INetCache\Content.Outlook",
    "$env:USERPROFILE\Documents\Outlook Files"
)

$cleanedFiles = 0
foreach ($path in $outlookPaths) {
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
    "C:\Windows\Temp\Outlook*",
    "$env:LOCALAPPDATA\Temp\Outlook*"
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
Write-Host "Outlook removal completed!" -ForegroundColor Green
Write-Host "Summary:" -ForegroundColor White
Write-Host "   Outlook processes terminated" -ForegroundColor Green
Write-Host "   Outlook services stopped and disabled" -ForegroundColor Green
Write-Host "   Outlook uninstalled" -ForegroundColor Green
Write-Host "   Registry entries cleaned" -ForegroundColor Green
Write-Host "   File system cleaned" -ForegroundColor Green
Write-Host ""
Write-Host "Please restart your computer to complete the removal process." -ForegroundColor Yellow
Write-Host "You can now install Outlook from the Windows Store." -ForegroundColor Cyan

Write-Host ""
Write-Host "Press Enter to exit..."
Read-Host
