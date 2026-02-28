import os
import sys
import time
import requests
import subprocess
import tempfile
from PyQt6.QtCore import QThread, pyqtSignal

# Update this to match the tag name on your GitHub Release (e.g., "v1.0.1")
CURRENT_VERSION = "v0.9.0"
GITHUB_REPO = "ibrahim-parvez/MRSI_Data_Normalization_Tool"

class AutoUpdater(QThread):
    check_finished = pyqtSignal(bool, str, str) # has_update, latest_version, download_url
    progress_updated = pyqtSignal(int)
    download_finished = pyqtSignal(str) # path to downloaded file
    error_occurred = pyqtSignal(str)

    def __init__(self, mode="check", url=""):
        super().__init__()
        self.mode = mode
        self.url = url

    def run(self):
        if self.mode == "check":
            self.check_for_updates()
        elif self.mode == "download":
            self.download_update()

    def check_for_updates(self):
        try:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            response = requests.get(url, timeout=5)
            response.raise_for_status()
            data = response.json()
            latest_version = data["tag_name"]
            
            if latest_version != CURRENT_VERSION:
                download_url = ""
                # Find the correct asset based on OS
                for asset in data["assets"]:
                    if sys.platform == "win32" and asset["name"].endswith(".exe"):
                        download_url = asset["browser_download_url"]
                    elif sys.platform == "darwin" and asset["name"].endswith(".zip"):
                        download_url = asset["browser_download_url"]
                
                if download_url:
                    self.check_finished.emit(True, latest_version, download_url)
                    return
            
            self.check_finished.emit(False, "", "")
        except Exception as e:
            self.error_occurred.emit(f"Failed to check for updates: {str(e)}")
            self.check_finished.emit(False, "", "")

    def download_update(self):
        try:
            response = requests.get(self.url, stream=True, timeout=10)
            response.raise_for_status()
            total_size = int(response.headers.get('content-length', 0))
            
            # FIX: Save the temp file to the system's temporary directory, outside the app
            temp_dir = tempfile.gettempdir()
            filename = "MRSI_Update_New.exe" if sys.platform == "win32" else "MRSI_Update_New.zip"
            download_path = os.path.join(temp_dir, filename)

            downloaded_size = 0
            with open(download_path, 'wb') as file:
                for data in response.iter_content(chunk_size=8192):
                    downloaded_size += len(data)
                    file.write(data)
                    if total_size > 0:
                        progress = int((downloaded_size / total_size) * 100)
                        self.progress_updated.emit(progress)

            self.download_finished.emit(download_path)
        except Exception as e:
            self.error_occurred.emit(f"Download failed: {str(e)}")
            self.download_finished.emit("")

def apply_update_and_restart(download_path):
    """Generates and runs an OS-specific script to replace the running executable."""
    is_frozen = getattr(sys, 'frozen', False)
    if not is_frozen:
        print("Cannot auto-update while running from raw Python script. Please compile first.")
        return

    current_exe = sys.executable
    
    # FIX: Save the scripts to the system temp directory so they survive the app deletion
    temp_dir = tempfile.gettempdir()

    if sys.platform == "win32":
        bat_path = os.path.join(temp_dir, "update_script.bat")
        with open(bat_path, "w") as bat_file:
            bat_file.write(f"""@echo off
                            timeout /t 2 /nobreak > NUL
                            move /y "{download_path}" "{current_exe}"
                            start "" "{current_exe}"
                            del "%~f0"
                            """)
        subprocess.Popen([bat_path], creationflags=subprocess.CREATE_NO_WINDOW)
        sys.exit(0)

    elif sys.platform == "darwin":
        # MAC: We are inside /Contents/MacOS/. We need to go up to the .app level.
        app_bundle_path = os.path.dirname(os.path.dirname(os.path.dirname(current_exe)))
        parent_dir = os.path.dirname(app_bundle_path) 
        app_name = os.path.basename(app_bundle_path)  

        sh_path = os.path.join(temp_dir, "update_script.sh")
        with open(sh_path, "w") as sh_file:
            sh_file.write(f"""#!/bin/bash
                            sleep 2
                            rm -rf "{app_bundle_path}"
                            unzip -q "{download_path}" -d "{parent_dir}" -x "__MACOSX/*"
                            rm "{download_path}"
                            open "{parent_dir}/{app_name}"
                            rm "$0"
                            """)
        os.chmod(sh_path, 0o755) 
        subprocess.Popen([sh_path], start_new_session=True)
        sys.exit(0)