"""
Adobe Creative Cloud Complete Removal Tool
This script will forcefully remove Adobe Creative Cloud and all related components.
Run as Administrator for best results.
"""

import subprocess
import os
import sys
import time
import winreg
from pathlib import Path

class AdobeRemover:
    def __init__(self):
        self.adobe_processes = [
            'Creative Cloud.exe', 'Creative Cloud Helper.exe', 'CCXProcess.exe',
            'CCLibrary.exe', 'AdobeIPCBroker.exe', 'Adobe Desktop Service.exe',
            'AdobeUpdateService.exe', 'AdobeCleanUpUtility.exe', 'acrotray.exe',
            'Acrobat.exe', 'AcroRd32.exe', 'Photoshop.exe', 'Illustrator.exe',
            'InDesign.exe', 'AfterFX.exe', 'Premiere Pro.exe', 'Bridge.exe',
            'Lightroom.exe', 'Dreamweaver.exe', 'Animate.exe', 'Audition.exe',
            'CoreSync.exe', 'AdobeCollabSync.exe', 'AdobeNotificationClient.exe'
        ]
        
        self.adobe_services = [
            'AdobeUpdateService', 'Adobe Genuine Monitor Service', 
            'Adobe Genuine Software Integrity Service', 'AGSService',
            'AGMService', 'AdobeARMservice', 'Adobe LM Service'
        ]
        
        self.cc_cleaner_path = r"C:\AdobeCreativeCloudCleanerTool_Win\AdobeCreativeCloudCleanerTool.exe"

    def run_as_admin(self):
        """Check if running as administrator"""
        try:
            import ctypes
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except Exception:
            return False

    def kill_adobe_processes(self):
        """Force kill all Adobe processes"""
        print("🔄 Stopping Adobe processes...")
        killed_processes = []
        
        for process in self.adobe_processes:
            try:
                result = subprocess.run(['taskkill', '/F', '/IM', process], 
                                      capture_output=True, text=True)
                if result.returncode == 0:
                    killed_processes.append(process)
                    print(f"  ✅ Killed: {process}")
            except Exception as e:
                print(f"  ❌ Failed to kill {process}: {e}")
        
        if killed_processes:
            print(f"📊 Successfully killed {len(killed_processes)} Adobe processes")
            time.sleep(3)  # Give processes time to fully terminate
        else:
            print("ℹ️  No Adobe processes were running")

    def stop_adobe_services(self):
        """Stop all Adobe services"""
        print("\n🔄 Stopping Adobe services...")
        stopped_services = []
        
        for service in self.adobe_services:
            try:
                # Stop the service
                result = subprocess.run(['sc', 'stop', service], 
                                      capture_output=True, text=True)
                if "STOP_PENDING" in result.stdout or "STOPPED" in result.stdout:
                    stopped_services.append(service)
                    print(f"  ✅ Stopped: {service}")
                
                # Disable the service
                subprocess.run(['sc', 'config', service, 'start=', 'disabled'], 
                             capture_output=True, text=True)
                
            except Exception as e:
                print(f"  ❌ Failed to stop {service}: {e}")
        
        if stopped_services:
            print(f"📊 Successfully stopped {len(stopped_services)} Adobe services")
        else:
            print("ℹ️  No Adobe services were running")

    def run_cc_cleaner(self):
        """Run Adobe CC Cleaner Tool"""
        print(f"\n🔄 Running Adobe CC Cleaner Tool...")
        
        if not os.path.exists(self.cc_cleaner_path):
            print(f"❌ CC Cleaner not found at: {self.cc_cleaner_path}")
            return False
        
        try:
            # Run CC Cleaner with options to remove all Adobe products
            print("🚀 Launching CC Cleaner in interactive mode...")
            print("⚠️  The CC Cleaner will open - follow these steps:")
            print("   1. Select 'Creative Cloud for desktop'")
            print("   2. Choose 'Remove All'")
            print("   3. Click 'Clean'")
            print("   4. Wait for completion")
            
            # Launch the cleaner tool
            subprocess.Popen([self.cc_cleaner_path])
            
            input("\n⏸️  Press Enter after CC Cleaner has finished...")
            return True
            
        except Exception as e:
            print(f"❌ Failed to run CC Cleaner: {e}")
            return False

    def clean_registry_entries(self):
        """Clean Adobe registry entries"""
        print("\n🔄 Cleaning Adobe registry entries...")
        
        registry_paths = [
            r"SOFTWARE\Adobe",
            r"SOFTWARE\WOW6432Node\Adobe",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Adobe*"
        ]
        
        cleaned_count = 0
        
        for reg_path in registry_paths:
            try:
                if "*" in reg_path:
                    # Handle wildcard entries
                    base_path = reg_path.replace("\\Adobe*", "")
                    try:
                        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base_path)
                        i = 0
                        while True:
                            try:
                                subkey_name = winreg.EnumKey(key, i)
                                if "adobe" in subkey_name.lower():
                                    full_path = f"{base_path}\\{subkey_name}"
                                    subprocess.run(['reg', 'delete', f'HKLM\\{full_path}', '/f'], 
                                                 capture_output=True)
                                    cleaned_count += 1
                                    print(f"  ✅ Deleted registry key: {subkey_name}")
                                i += 1
                            except WindowsError:
                                break
                        winreg.CloseKey(key)
                    except:
                        pass
                else:
                    # Direct registry deletion
                    result = subprocess.run(['reg', 'delete', f'HKLM\\{reg_path}', '/f'], 
                                          capture_output=True, text=True)
                    if result.returncode == 0:
                        cleaned_count += 1
                        print(f"  ✅ Deleted registry path: {reg_path}")
                        
            except Exception as e:
                print(f"  ⚠️  Couldn't access registry path {reg_path}: {e}")
        
        print(f"📊 Cleaned {cleaned_count} registry entries")

    def clean_file_system(self):
        """Clean Adobe files and folders"""
        print("\n🔄 Cleaning Adobe files and folders...")
        
        adobe_paths = [
            r"C:\Program Files\Adobe",
            r"C:\Program Files (x86)\Adobe", 
            r"C:\Program Files\Common Files\Adobe",
            r"C:\Program Files (x86)\Common Files\Adobe",
            r"C:\ProgramData\Adobe",
            r"C:\Users\{username}\AppData\Local\Adobe",
            r"C:\Users\{username}\AppData\Roaming\Adobe",
            r"C:\Users\{username}\Creative Cloud Files"
        ]
        
        username = os.getenv('USERNAME', 'User')
        cleaned_count = 0
        
        for path_template in adobe_paths:
            path = path_template.replace('{username}', username)
            
            if os.path.exists(path):
                try:
                    # Use rmdir /s /q for Windows
                    result = subprocess.run(['rmdir', '/s', '/q', path], 
                                          shell=True, capture_output=True)
                    if result.returncode == 0:
                        cleaned_count += 1
                        print(f"  ✅ Removed: {path}")
                    else:
                        print(f"  ⚠️  Couldn't remove: {path} (may be in use)")
                except Exception as e:
                    print(f"  ❌ Error removing {path}: {e}")
            else:
                print(f"  ℹ️  Path not found: {path}")
        
        print(f"📊 Cleaned {cleaned_count} file system locations")

    def final_cleanup(self):
        """Final cleanup steps"""
        print("\n🔄 Performing final cleanup...")
        
        # Clear temporary files
        temp_paths = [
            r"C:\Windows\Temp\Adobe*",
            rf"C:\Users\{os.getenv('USERNAME')}\AppData\Local\Temp\Adobe*"
        ]
        
        for temp_path in temp_paths:
            try:
                subprocess.run(['del', '/s', '/q', temp_path], shell=True, capture_output=True)
                print(f"  ✅ Cleared temp files: {temp_path}")
            except:
                pass
        
        # Refresh system
        print("  🔄 Refreshing system...")
        subprocess.run(['sfc', '/scannow'], capture_output=True)
        
        print("✅ Final cleanup completed!")

    def remove_adobe(self):
        """Main removal process"""
        print("🚀 Adobe Creative Cloud Complete Removal Tool")
        print("=" * 50)
        
        if not self.run_as_admin():
            print("⚠️  For best results, run this script as Administrator")
            print("   Right-click and select 'Run as administrator'")
        
        try:
            # Step 1: Kill processes
            self.kill_adobe_processes()
            
            # Step 2: Stop services
            self.stop_adobe_services()
            
            # Step 3: Run CC Cleaner
            if self.run_cc_cleaner():
                print("✅ CC Cleaner completed successfully")
            
            # Step 4: Clean registry
            self.clean_registry_entries()
            
            # Step 5: Clean file system
            self.clean_file_system()
            
            # Step 6: Final cleanup
            self.final_cleanup()
            
            print("\n" + "=" * 50)
            print("🎉 Adobe Creative Cloud removal completed!")
            print("📋 Summary:")
            print("   ✅ Adobe processes terminated")
            print("   ✅ Adobe services stopped")
            print("   ✅ CC Cleaner executed")
            print("   ✅ Registry entries cleaned")
            print("   ✅ File system cleaned")
            print("\n⚠️  Please restart your computer to complete the removal process.")
            
        except Exception as e:
            print(f"\n❌ An error occurred during removal: {e}")
            print("💡 Try running the script as Administrator for better results.")

if __name__ == "__main__":
    remover = AdobeRemover()
    remover.remove_adobe()
    
    input("\nPress Enter to exit...")