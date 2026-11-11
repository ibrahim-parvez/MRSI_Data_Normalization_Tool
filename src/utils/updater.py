"""Checking GitHub for a newer release, downloading it, and swapping it in.

Releases are published by the Release workflow (.github/workflows/release.yml).
Each one carries exactly two app files, and this module looks for them by name:

    Windows  MRSI.Data.Normalization.Tool.exe       (single-file build)
    macOS    MRSI.Data.Normalization.Tool.app.zip   (the .app, zipped with ditto)

The separate "Installer" release is never treated as a version.
"""

import hashlib
import os
import sys
import tempfile
import subprocess

import requests
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (QDialog, QGridLayout, QHBoxLayout, QLabel, QPushButton,
                             QTextEdit, QVBoxLayout)

from version import __version__, parse_version

CURRENT_VERSION = f"v{__version__}"
GITHUB_REPO = "ibrahim-parvez/MRSI_Data_Normalization_Tool"

APP_ASSET_NAMES = {
    "win32": "MRSI.Data.Normalization.Tool.exe",
    "darwin": "MRSI.Data.Normalization.Tool.app.zip",
}

# Filled by a check, read by the download and the dialog: what GitHub said
# about each release file (size, checksum, notes), keyed by download URL.
_RELEASE_INFO = {}


def is_installed_build():
    """True for the packaged app. Running from source cannot update itself."""
    return bool(getattr(sys, "frozen", False))


def release_notes(download_url):
    return _RELEASE_INFO.get(download_url, {}).get("notes", "")


def _pick_app_asset(assets):
    """The app file for this platform, or None. Matched by exact name first,
    then by shape, and never an installer."""
    wanted = APP_ASSET_NAMES.get(sys.platform)
    if not wanted:
        return None
    for asset in assets:
        if asset.get("name") == wanted:
            return asset
    suffix = ".exe" if sys.platform == "win32" else ".app.zip"
    for asset in assets:
        name = asset.get("name", "").lower()
        if name.endswith(suffix) and "installer" not in name:
            return asset
    return None


class UpdateAvailableDialog(QDialog):
    """Asks whether to update, showing what changed."""

    def __init__(self, current_version, new_version, parent=None, notes=""):
        super().__init__(parent)
        self.setWindowTitle("Update Available")
        self.setMinimumWidth(380)

        self.dark_mode = getattr(parent, 'dark_mode', False)
        if self.dark_mode:
            self.setStyleSheet("QDialog { background-color: #2A2B2E; color: #E8EAED; } QLabel { color: #E8EAED; }"
                               " QTextEdit { background-color: #1E1F22; color: #E8EAED; border: 1px solid #444; border-radius: 4px; }")
            header_color = "#FF6B8B"  # lighter maroon, readable on dark
            cancel_bg, cancel_hover = "#424242", "#616161"
        else:
            self.setStyleSheet("QDialog { background-color: #FAFAFA; color: #202124; } QLabel { color: #202124; }"
                               " QTextEdit { background-color: #FFFFFF; color: #202124; border: 1px solid #DDD; border-radius: 4px; }")
            header_color = "#7A003C"  # McMaster maroon
            cancel_bg, cancel_hover = "#9E9E9E", "#757575"

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 16)

        header_label = QLabel("A new version is available")
        header_label.setStyleSheet(f"font-size: 16px; font-weight: bold; color: {header_color};")
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header_label)
        layout.addSpacing(10)

        version_layout = QGridLayout()
        version_layout.setContentsMargins(20, 0, 20, 0)
        version_layout.addWidget(QLabel("<b>Current version:</b>"), 0, 0)
        version_layout.addWidget(QLabel(current_version), 0, 1)
        version_layout.addWidget(QLabel("<b>New version:</b>"), 1, 0)
        version_layout.addWidget(QLabel(f"<span style='color: #4CAF50; font-weight: bold;'>{new_version}</span>"), 1, 1)
        layout.addLayout(version_layout)

        if notes and notes.strip():
            layout.addSpacing(8)
            layout.addWidget(QLabel("<b>What's new</b>"))
            notes_box = QTextEdit()
            notes_box.setReadOnly(True)
            notes_box.setMarkdown(notes)
            notes_box.setMaximumHeight(150)
            layout.addWidget(notes_box)

        layout.addSpacing(10)
        btn_layout = QHBoxLayout()
        self.later_btn = QPushButton("Update Later")
        self.update_btn = QPushButton("Update Now")
        self.update_btn.setStyleSheet("""
            QPushButton { background-color: #4CAF50; color: white; padding: 8px 15px; font-weight: bold; border-radius: 4px; border: none; }
            QPushButton:hover { background-color: #45A049; }
        """)
        self.later_btn.setStyleSheet(f"""
            QPushButton {{ background-color: {cancel_bg}; color: white; padding: 8px 15px; border-radius: 4px; border: none; }}
            QPushButton:hover {{ background-color: {cancel_hover}; }}
        """)
        for btn in (self.update_btn, self.later_btn):
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.update_btn.setDefault(True)
        self.update_btn.clicked.connect(self.accept)
        self.later_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self.later_btn)
        btn_layout.addWidget(self.update_btn)
        layout.addLayout(btn_layout)


class AutoUpdater(QThread):
    check_finished = pyqtSignal(bool, str, str)  # has_update, latest_version, download_url
    progress_updated = pyqtSignal(int)
    download_finished = pyqtSignal(str)          # path to the verified download
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
        """Find the highest released version that has a file for this
        platform, and report it only if it is newer than this build."""
        try:
            response = requests.get(
                f"https://api.github.com/repos/{GITHUB_REPO}/releases",
                params={"per_page": 50},
                headers={"Accept": "application/vnd.github+json",
                         "User-Agent": "MRSI-DNT-Updater"},
                timeout=8,
            )
            response.raise_for_status()

            best = None
            for release in response.json():
                if release.get("draft") or release.get("prerelease"):
                    continue
                version = parse_version(release.get("tag_name"))
                asset = _pick_app_asset(release.get("assets", []))
                if version and asset and (best is None or version > best[0]):
                    best = (version, release, asset)

            current = parse_version(CURRENT_VERSION)
            if best and current and best[0] > current:
                _, release, asset = best
                url = asset["browser_download_url"]
                digest = asset.get("digest") or ""
                _RELEASE_INFO[url] = {
                    "size": asset.get("size") or 0,
                    "sha256": digest.split(":", 1)[1] if digest.startswith("sha256:") else "",
                    "notes": release.get("body") or "",
                }
                self.check_finished.emit(True, release["tag_name"], url)
            else:
                self.check_finished.emit(False, "", "")
        except Exception as e:
            self.error_occurred.emit(f"Failed to check for updates: {e}")

    def download_update(self):
        """Download to a temporary name, check size and checksum, and only
        then hand the file over. A cut-off download is never installed."""
        info = _RELEASE_INFO.get(self.url, {})
        filename = "MRSI_Update_New.exe" if sys.platform == "win32" else "MRSI_Update_New.zip"
        final_path = os.path.join(tempfile.gettempdir(), filename)
        part_path = final_path + ".part"
        try:
            response = requests.get(self.url, stream=True, timeout=30,
                                    headers={"User-Agent": "MRSI-DNT-Updater"})
            response.raise_for_status()
            total_size = int(response.headers.get('content-length', 0)) or info.get("size", 0)

            sha = hashlib.sha256()
            downloaded = 0
            with open(part_path, 'wb') as file:
                for chunk in response.iter_content(chunk_size=1024 * 256):
                    file.write(chunk)
                    sha.update(chunk)
                    downloaded += len(chunk)
                    if total_size > 0:
                        self.progress_updated.emit(min(100, int(downloaded * 100 / total_size)))

            expected_size = info.get("size") or total_size
            if expected_size and downloaded != expected_size:
                raise IOError(f"the download was incomplete ({downloaded:,} of {expected_size:,} bytes)")
            expected_sha = info.get("sha256")
            if expected_sha and sha.hexdigest() != expected_sha:
                raise IOError("the downloaded file did not match the release checksum")

            os.replace(part_path, final_path)
            self.download_finished.emit(final_path)
        except Exception as e:
            try:
                os.remove(part_path)
            except OSError:
                pass
            self.error_occurred.emit(f"Download failed: {e}")

# The update has to be applied by something that outlives this process, since
# the app cannot overwrite itself while it is running. Both scripts below
# follow the same rule: never destroy the installed copy until the replacement
# is unpacked and verified, and always leave the user with a working app.
#
# The failure both scripts are guarding against is a permission one. The
# installer puts the app wherever the user pointed it - usually ~/Desktop, but
# it can be anywhere - and that location may not be writable by the time an
# update lands: /Applications or Program Files for a non-admin account, a
# synced folder, a locked-down managed machine.

_WINDOWS_UPDATE_SCRIPT = r"""@echo off
setlocal enableextensions
set "TARGET={target}"
set "SOURCE={source}"
set /a TRIES=0

:retry
set /a TRIES+=1
ping 127.0.0.1 -n 3 > NUL
move /y "%SOURCE%" "%TARGET%" > NUL 2>&1
if not errorlevel 1 goto relaunch
rem The first failures are the old process still holding its file lock, which
rem clears on its own. A permission problem never clears, so give up rather
rem than spin forever with the app closed.
if %TRIES% lss 20 goto retry

rem Quoting a message through powershell from batch is fragile, and this path
rem only runs when something is already wrong; a text file always shows up.
set "REPORT=%TEMP%\MRSI_update_failed.txt"
break > "%REPORT%"
>> "%REPORT%" echo The update could not be installed.
>> "%REPORT%" echo.
>> "%REPORT%" echo This file could not be replaced:
>> "%REPORT%" echo   %TARGET%
>> "%REPORT%" echo.
>> "%REPORT%" echo That usually means the app sits in a folder needing administrator
>> "%REPORT%" echo rights, or another program is holding it open.
>> "%REPORT%" echo.
>> "%REPORT%" echo The downloaded update was kept here, so you can move it in by hand:
>> "%REPORT%" echo   %SOURCE%
>> "%REPORT%" echo.
>> "%REPORT%" echo Your existing version still works and will reopen now.
start "" notepad "%REPORT%"

:relaunch
start "" "%TARGET%"
del "%~f0"
"""

_MACOS_UPDATE_SCRIPT = """#!/bin/bash
# Deliberately not 'set -e': every failure routes through fail() so the
# installed app is always put back and reopened.
APP={app}
PARENT={parent}
NAME={name}
ZIP={zip}
STAGE="$(mktemp -d -t mrsi_update)"
BACKUP=""

sleep 2

fail() {{
    # Put the old app back if it has already been moved aside.
    if [ -n "$BACKUP" ] && [ -d "$BACKUP" ] && [ ! -d "$APP" ]; then
        mv "$BACKUP" "$APP"
    fi
    /usr/bin/osascript -e "display alert \"Update failed\" message \"$1
        
Your existing version still works and will reopen now.\" as critical" > /dev/null 2>&1
    rm -rf "$STAGE"
    open "$APP" > /dev/null 2>&1
    rm -f "$0"
    exit 1
}}

ditto -x -k "$ZIP" "$STAGE" || fail "The downloaded update could not be unpacked."
rm -rf "$STAGE/__MACOSX"

# The zip usually holds the .app at its root, but tolerate one extra folder.
NEW="$STAGE/$NAME"
if [ ! -d "$NEW" ]; then
    NEW="$(find "$STAGE" -maxdepth 2 -name "*.app" -print -quit)"
fi
[ -n "$NEW" ] && [ -d "$NEW" ] || fail "The downloaded update did not contain an application."

[ -w "$PARENT" ] || fail "\"$PARENT\" is not writable, so the update could not be installed there."

# Swap by moving, not by deleting first: if anything goes wrong the installed
# copy is still sitting in $BACKUP and fail() restores it.
BACKUP="$PARENT/.$NAME.previous.$$"
mv "$APP" "$BACKUP" || fail "The installed app could not be replaced."
if ! mv "$NEW" "$APP"; then
    fail "The updated app could not be moved into place."
fi

# A zip that came through a browser carries a quarantine flag that would make
# the fresh copy look untrusted on first launch.
xattr -dr com.apple.quarantine "$APP" > /dev/null 2>&1

rm -rf "$BACKUP" "$STAGE"
rm -f "$ZIP"
open "$APP"
rm -f "$0"
"""


def _shell_quote(value):
    """Single-quote `value` for safe interpolation into the bash script."""
    return "'" + str(value).replace("'", "'\\''") + "'"


def apply_update_and_restart(download_path):
    """Hand the download to an OS-specific script that swaps it in once this
    process exits, then exit. Returns False (and does nothing) when the app
    cannot update itself, e.g. when running from source."""
    if not is_installed_build():
        print("Cannot auto-update while running from source. Build the app first.")
        return False

    current_exe = sys.executable
    temp_dir = tempfile.gettempdir()

    # PyInstaller's own variables confuse a child process that is itself a
    # frozen app; strip them so the relaunched copy starts clean.
    clean_env = os.environ.copy()
    for key in [k for k in clean_env if k.startswith('_PYI_') or k.startswith('_MEI')]:
        clean_env.pop(key, None)

    if sys.platform == "win32":
        bat_path = os.path.join(temp_dir, "update_script.bat")
        with open(bat_path, "w") as bat_file:
            bat_file.write(
                _WINDOWS_UPDATE_SCRIPT.format(target=current_exe, source=download_path)
            )

        subprocess.Popen(
            [bat_path],
            creationflags=subprocess.CREATE_NO_WINDOW,
            env=clean_env,
        )
        # Exit immediately so the .exe lock is released for the updater
        os._exit(0)

    elif sys.platform == "darwin":
        # sys.executable is <bundle>.app/Contents/MacOS/<exe>.
        app_bundle_path = os.path.dirname(os.path.dirname(os.path.dirname(current_exe)))
        parent_dir = os.path.dirname(app_bundle_path)
        app_name = os.path.basename(app_bundle_path)

        if not app_bundle_path.endswith(".app"):
            print("Not running from a .app bundle - refusing to self-update.")
            return False

        sh_path = os.path.join(temp_dir, "update_script.sh")
        with open(sh_path, "w") as sh_file:
            sh_file.write(_MACOS_UPDATE_SCRIPT.format(
                app=_shell_quote(app_bundle_path),
                parent=_shell_quote(parent_dir),
                name=_shell_quote(app_name),
                zip=_shell_quote(download_path),
            ))
        os.chmod(sh_path, 0o755)

        subprocess.Popen([sh_path], start_new_session=True, env=clean_env)
        # Same immediate exit on Mac
        os._exit(0)

    return False  # unsupported platform
