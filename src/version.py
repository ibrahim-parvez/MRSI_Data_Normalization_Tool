"""The app's version, in one place.

The Release workflow (.github/workflows/release.yml) rewrites ``__version__``
for each release, so there is normally no need to edit it by hand. The
updater, the installer and the About box all read it from here.

Kept free of imports so the installer can use it without pulling in Qt.
"""

import re

__version__ = "2.1.1"

_VERSION_RE = re.compile(r"v?(\d+)\.(\d+)\.(\d+)")


def parse_version(text):
    """'v2.1.1' or '2.1.1' -> (2, 1, 1). None for anything else, such as the
    'Installer' release tag, so it can never be mistaken for a version."""
    match = _VERSION_RE.fullmatch(str(text or "").strip())
    return tuple(int(part) for part in match.groups()) if match else None
