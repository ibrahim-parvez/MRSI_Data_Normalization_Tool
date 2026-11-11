# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the standalone MRSI DNT installer.

Build it with ``python packaging/build.py installer``. This is the small
app end users download first; it fetches a release from GitHub and unpacks
the tool itself. Single file on both platforms, since there is no in-place
updater to worry about.
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(SPECPATH).parent

# Shown in Finder's Get Info. src/version.py is the single source; the
# release workflow stamps it before building.
import re
APP_VERSION = re.search(
    r'^__version__ = "([^"]+)"',
    (PROJECT_ROOT / "src" / "version.py").read_text(encoding="utf-8"),
    re.MULTILINE,
).group(1)
ASSETS = PROJECT_ROOT / "assets" / "images"

IS_MACOS = sys.platform == "darwin"

APP_NAME = "MRSI_DNT_Installer"
BUNDLE_ID = "ca.mcmaster.mrsi.dnt.installer"

icon = str(ASSETS / ("logo.icns" if IS_MACOS else "logo.ico"))

a = Analysis(
    [str(PROJECT_ROOT / "src" / "utils" / "installer" / "installer.py")],
    pathex=[str(PROJECT_ROOT / "src")],
    binaries=[],
    datas=[(str(ASSETS), "assets/images")],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=(
        str(PROJECT_ROOT / "packaging" / "entitlements.plist") if IS_MACOS else None
    ),
    icon=icon,
)

if IS_MACOS:
    # The installer unpacks the tool into a folder the user picks, and offers
    # ~/Desktop by default. Desktop, Documents and Downloads are all behind
    # TCC, and macOS only prompts for them if the bundle says why it wants in
    # - without these strings the write is refused and the install half-fails.
    info_plist = {
        "NSDesktopFolderUsageDescription": (
            "The installer needs access to install the MRSI Data Normalization "
            "Tool into the folder you choose on your Desktop."
        ),
        "NSDocumentsFolderUsageDescription": (
            "The installer needs access to install the MRSI Data Normalization "
            "Tool into the folder you choose in Documents."
        ),
        "NSDownloadsFolderUsageDescription": (
            "The installer needs access to install the MRSI Data Normalization "
            "Tool into the folder you choose in Downloads."
        ),
        "NSHighResolutionCapable": True,
    }

    app = BUNDLE(
        exe,
        name=f"{APP_NAME}.app",
        icon=icon,
        bundle_identifier=BUNDLE_ID,
        version=APP_VERSION,
        info_plist=info_plist,
    )
