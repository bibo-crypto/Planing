"""sync_venv_packages.py — prune packages from the active virtual
environment that are no longer listed in requirements.txt.

Run with the venv's own Python (build.bat does this after activating
venv, before the normal `pip install -r requirements.txt` step). Adding
packages is already something `pip install -r requirements.txt` handles
correctly on its own (it skips anything already at the right version);
the one thing it never does is *remove* a package that used to be in
requirements.txt but isn't anymore. This script closes that gap so
build.bat can update an existing venv in place instead of deleting and
recreating it from scratch on every build.
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

# Never remove these even if requirements.txt doesn't list them --
# pip/setuptools/wheel are the venv's own bootstrap tooling, not an
# app dependency, and PyInstaller is invoked by build.bat itself
# separately from requirements.txt.
ALWAYS_KEEP = {"pip", "setuptools", "wheel"}


def _normalize(name: str) -> str:
    """PEP 503 style: case-insensitive, '-'/'_'/'.' treated the same."""
    return re.sub(r"[-_.]+", "-", name.strip().lower())


def _requirement_names(requirements_path: Path) -> set[str]:
    names = set()
    if not requirements_path.is_file():
        return names
    for line in requirements_path.read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].strip()
        if not line or line.startswith(("-", "git+", "http://", "https://")):
            continue
        # Package name is everything before the first version/extra marker.
        match = re.match(r"^([A-Za-z0-9][A-Za-z0-9._-]*)", line)
        if match:
            names.add(_normalize(match.group(1)))
    return names


def _installed_names() -> set[str]:
    result = subprocess.run(
        [sys.executable, "-m", "pip", "list", "--format=freeze"],
        capture_output=True, text=True, check=True,
    )
    names = set()
    for line in result.stdout.splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "==" not in line:
            continue
        names.add(_normalize(line.split("==", 1)[0]))
    return names


def main() -> int:
    requirements_path = Path(__file__).resolve().parent / "requirements.txt"
    wanted = _requirement_names(requirements_path) | {_normalize(p) for p in ALWAYS_KEEP}
    installed = _installed_names()
    to_remove = sorted(installed - wanted)

    if not to_remove:
        print("[sync_venv] Nothing to remove -- venv already matches requirements.txt.")
        return 0

    print(f"[sync_venv] Removing {len(to_remove)} package(s) no longer in requirements.txt:")
    for name in to_remove:
        print(f"  - {name}")

    subprocess.run(
        [sys.executable, "-m", "pip", "uninstall", "-y", "--quiet", *to_remove],
        check=False,  # a single stubborn package shouldn't abort the whole build
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
