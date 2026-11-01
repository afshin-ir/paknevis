#!/usr/bin/env python3
"""Build the LibreOffice extension package from the repository sources."""

from pathlib import Path
import sys
import zipfile

ROOT = Path(__file__).resolve().parent
OUTPUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "Paknevis.oxt"
FILES = (
    ("Addons.xcu", ROOT / "extension" / "Addons.xcu"),
    ("description.xml", ROOT / "extension" / "description.xml"),
    ("META-INF/manifest.xml", ROOT / "extension" / "META-INF" / "manifest.xml"),
    ("registration/license.txt", ROOT / "extension" / "registration" / "license.txt"),
    ("PAKNEVIS.py", ROOT / "PAKNEVIS.py"),
    ("DocumentList.json", ROOT / "DocumentList.json"),
)

with zipfile.ZipFile(OUTPUT, "w", compression=zipfile.ZIP_DEFLATED) as package:
    for archive_path, source_path in FILES:
        package.write(source_path, archive_path)

print(f"Built {OUTPUT}")
