"""Release helpers used by .github/workflows/release.yml.

Each subcommand is one small step of a release, kept here (rather than as
shell inside the workflow) so it behaves the same on Windows and macOS
runners and can be tried locally:

    python packaging/release.py next patch       # print the next version, e.g. 2.1.2
    python packaging/release.py set 2.1.2        # write it into src/version.py
    python packaging/release.py smoke            # launch the built app, check it starts
    python packaging/release.py smoke --source   # same check, straight from src/
    python packaging/release.py package          # dist/ -> dist/release/ with release names
    python packaging/release.py notes            # release notes from $RELEASE_NOTES

Run `python packaging/build.py all` before `smoke` and `package`.
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DIST = PROJECT_ROOT / "dist"
RELEASE_DIR = DIST / "release"
VERSION_FILE = PROJECT_ROOT / "src" / "version.py"

APP_NAME = "MRSI Data Normalization Tool"
INSTALLER_NAME = "MRSI_DNT_Installer"

# The names the updater and installer look for. Changing these breaks
# updating for everyone already installed, so leave them alone.
APP_ASSET = {"win32": "MRSI.Data.Normalization.Tool.exe",
             "darwin": "MRSI.Data.Normalization.Tool.app.zip"}
INSTALLER_ASSET = {"win32": "MRSI_DNT_Installer.exe",
                   "darwin": "MRSI_DNT_Installer.app.zip"}

VERSION_RE = re.compile(r"^v(\d+)\.(\d+)\.(\d+)$")


def latest_tag_version():
    """Highest vX.Y.Z tag in the repository, as a tuple (0, 0, 0 if none)."""
    tags = subprocess.run(["git", "tag", "--list", "v*"], cwd=PROJECT_ROOT,
                          capture_output=True, text=True, check=True).stdout.split()
    versions = [tuple(map(int, m.groups())) for t in tags if (m := VERSION_RE.match(t))]
    return max(versions, default=(0, 0, 0))


def cmd_next(args):
    major, minor, patch = latest_tag_version()
    if args.bump == "major":
        major, minor, patch = major + 1, 0, 0
    elif args.bump == "minor":
        minor, patch = minor + 1, 0
    else:
        patch += 1
    print(f"{major}.{minor}.{patch}")


def cmd_set(args):
    if not re.fullmatch(r"\d+\.\d+\.\d+", args.version):
        sys.exit(f"Not a version: {args.version!r} (expected e.g. 2.1.2)")
    text = VERSION_FILE.read_text(encoding="utf-8")
    new_text, count = re.subn(r'^__version__ = "[^"]*"$', f'__version__ = "{args.version}"',
                              text, flags=re.MULTILINE)
    if count != 1:
        sys.exit(f"Could not find the __version__ line in {VERSION_FILE}")
    VERSION_FILE.write_text(new_text, encoding="utf-8")
    print(f"src/version.py -> {args.version}")


def _app_executable():
    if sys.platform == "win32":
        return DIST / f"{APP_NAME}.exe"
    return DIST / f"{APP_NAME}.app" / "Contents" / "MacOS" / APP_NAME


def cmd_smoke(args):
    """Start the app in smoke-test mode: it opens the main window and quits on
    its own. A non-zero exit or a hang fails the check. Qt draws offscreen, so
    the result does not depend on the build machine having a display."""
    if args.source:
        command = [sys.executable, str(PROJECT_ROOT / "src" / "main.py")]
    else:
        exe = _app_executable()
        if not exe.exists():
            sys.exit(f"Built app not found: {exe}")
        command = [str(exe)]
    env = dict(os.environ, MRSI_SMOKE_TEST="1", QT_QPA_PLATFORM="offscreen")
    try:
        result = subprocess.run(command, env=env, timeout=args.timeout)
    except subprocess.TimeoutExpired:
        sys.exit(f"Smoke test failed: the app did not finish starting within {args.timeout}s")
    if result.returncode != 0:
        sys.exit(f"Smoke test failed: the app exited with code {result.returncode}")
    print("Smoke test passed: the app starts and opens its main window.")


def _zip_app(app_path, zip_path):
    # ditto, as Finder's "Compress" does: it keeps the symlinks inside the
    # Qt frameworks, which `zip -r` would turn into copies and break.
    subprocess.run(["ditto", "-c", "-k", "--sequesterRsrc", "--keepParent",
                    str(app_path), str(zip_path)], check=True)


def cmd_package(args):
    if RELEASE_DIR.exists():
        shutil.rmtree(RELEASE_DIR)
    RELEASE_DIR.mkdir(parents=True)
    platform = "win32" if sys.platform == "win32" else "darwin"

    if platform == "win32":
        shutil.copy2(DIST / f"{APP_NAME}.exe", RELEASE_DIR / APP_ASSET[platform])
        shutil.copy2(DIST / f"{INSTALLER_NAME}.exe", RELEASE_DIR / INSTALLER_ASSET[platform])
    else:
        _zip_app(DIST / f"{APP_NAME}.app", RELEASE_DIR / APP_ASSET[platform])
        _zip_app(DIST / f"{INSTALLER_NAME}.app", RELEASE_DIR / INSTALLER_ASSET[platform])

    for item in sorted(RELEASE_DIR.iterdir()):
        print(f"{item.name}  ({item.stat().st_size / 1_048_576:.1f} MB)")


def cmd_notes(args):
    """Turn the one-line notes typed into the workflow form into a bullet
    list: items are separated with semicolons."""
    raw = os.environ.get("RELEASE_NOTES", "")
    items = [part.strip(" -\t") for part in raw.split(";") if part.strip(" -\t")]
    lines = [f"- {item}" for item in items] or ["- Bug fixes and improvements"]
    print("\n".join(lines))


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("next", help="print the next version number")
    p.add_argument("bump", choices=["patch", "minor", "major"])
    p.set_defaults(func=cmd_next)

    p = sub.add_parser("set", help="write a version into src/version.py")
    p.add_argument("version")
    p.set_defaults(func=cmd_set)

    p = sub.add_parser("smoke", help="check the app starts")
    p.add_argument("--source", action="store_true", help="run src/main.py instead of the build")
    p.add_argument("--timeout", type=int, default=180)
    p.set_defaults(func=cmd_smoke)

    p = sub.add_parser("package", help="copy/zip the build into dist/release")
    p.set_defaults(func=cmd_package)

    p = sub.add_parser("notes", help="format release notes from $RELEASE_NOTES")
    p.set_defaults(func=cmd_notes)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
