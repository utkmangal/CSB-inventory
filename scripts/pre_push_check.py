#!/usr/bin/env python3
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

required_files = [
    ROOT / "index.html",
    ROOT / "favicon.svg",
    ROOT / ".github" / "workflows" / "deploy.yml",
]

required_dirs = [
    ROOT / "images",
]

missing = [str(path.relative_to(ROOT)) for path in required_files if not path.exists()]
missing_dirs = [str(path.relative_to(ROOT)) for path in required_dirs if not path.exists()]

errors = []
if missing:
    errors.append("Missing required files: " + ", ".join(missing))

html_path = ROOT / "index.html"
if html_path.exists():
    html = html_path.read_text(encoding="utf-8")
    if "<!DOCTYPE html>" not in html:
        errors.append("index.html does not look like a valid HTML document")
    if "favicon.svg" not in html:
        errors.append("index.html is missing the favicon reference")

if errors:
    print("Pre-push integrity check failed:")
    for error in errors:
        print(f"- {error}")
    sys.exit(1)

print("Pre-push integrity check passed.")
print("Verified files:")
for path in required_files:
    print(f"- {path.relative_to(ROOT)}")
for path in required_dirs:
    print(f"- {path.relative_to(ROOT)}/")
