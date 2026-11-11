# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the MRSI Data Normalization Tool.

Build it with ``python packaging/build.py app`` rather than by hand, so the
work and dist directories always land in the same place.

clean dist -> rm -rf build dist && rm -f *.spec 

Windows produces a single self-contained .exe with a native splash image;
macOS produces a .app bundle, which is the layout the in-app updater expects
when it swaps a new version in.
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

IS_WINDOWS = sys.platform == "win32"
IS_MACOS = sys.platform == "darwin"

APP_NAME = "MRSI Data Normalization Tool"
BUNDLE_ID = "ca.mcmaster.mrsi.dnt"

# assets/images has to travel with the build: utils.resources reads the logo
# out of it at runtime via sys._MEIPASS.
datas = [(str(ASSETS), "assets/images")]

icon = str(ASSETS / ("logo.icns" if IS_MACOS else "logo.ico"))

a = Analysis(
    [str(PROJECT_ROOT / "src" / "main.py")],
    pathex=[str(PROJECT_ROOT / "src")],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

# The bootloader splash covers the gap before Qt is up; main.py closes it as
# soon as its own splash screen is painted. Windows-only, one-file builds only.
splash = (
    Splash(
        str(ASSETS / "mrsi_logo.png"),
        binaries=a.binaries,
        datas=a.datas,
        always_on_top=True,
    )
    if IS_WINDOWS
    else None
)

if IS_WINDOWS:
    # One file: everything is folded into the .exe the updater replaces.
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        splash,
        splash.binaries,
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
        entitlements_file=None,
        icon=icon,
    )
else:
    # One directory, then wrapped in a .app bundle.
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name=APP_NAME,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=str(PROJECT_ROOT / "packaging" / "entitlements.plist"),
        icon=icon,
    )

    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name=APP_NAME,
    )

    # macOS gates Apple Events behind TCC: the first time the app drives
    # Microsoft Excel (xlwings -> appscript), the system shows a consent
    # prompt - but only if the bundle declares why it wants the permission.
    # With no NSAppleEventsUsageDescription the prompt is skipped and every
    # event is refused outright, which surfaces from appscript as:
    #
    #     OSERROR: -1743
    #     MESSAGE: The user has declined permission.
    #
    # The folder strings cover the same gate for the locations the tool reads
    # and writes (Desktop is the default output for a Combine run).
    info_plist = {
        "NSAppleEventsUsageDescription": (
            f"{APP_NAME} controls Microsoft Excel to convert .xls files, "
            "recalculate formulas and save your workbooks."
        ),
        "NSDesktopFolderUsageDescription": (
            f"{APP_NAME} reads and writes normalization workbooks on your Desktop."
        ),
        "NSDocumentsFolderUsageDescription": (
            f"{APP_NAME} reads and writes normalization workbooks in your Documents folder."
        ),
        "NSDownloadsFolderUsageDescription": (
            f"{APP_NAME} reads and writes normalization workbooks in your Downloads folder."
        ),
        # Data for this tool often lives on a USB stick or a mounted lab
        # share, and both are gated separately from the home folders.
        "NSRemovableVolumesUsageDescription": (
            f"{APP_NAME} needs access to open and save workbooks on USB and "
            "other removable drives."
        ),
        "NSNetworkVolumesUsageDescription": (
            f"{APP_NAME} needs access to open and save workbooks on network "
            "drives and mounted shares."
        ),
        "NSHighResolutionCapable": True,
    }

    app = BUNDLE(
        coll,
        name=f"{APP_NAME}.app",
        icon=icon,
        bundle_identifier=BUNDLE_ID,
        version=APP_VERSION,
        info_plist=info_plist,
    )
