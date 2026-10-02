"""Portable backup/restore for Planning application state."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
import shutil
import tempfile
import zipfile

from utility.utils import APP_DATA_DIR


def create_backup(destination: str | Path) -> Path:
    """Zip application data, SQLite databases, settings and caches."""
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        if APP_DATA_DIR.exists():
            for path in APP_DATA_DIR.rglob("*"):
                if not path.is_file():
                    continue
                relative = path.relative_to(APP_DATA_DIR)
                if relative.parts and relative.parts[0] in {"backups", "logs"}:
                    continue
                archive.write(path, arcname=str(relative))
        archive.writestr("backup_info.txt", f"Created: {datetime.now().isoformat()}\n")
    return destination


def restore_backup(source: str | Path) -> int:
    """Safely restore a backup into APP_DATA_DIR; return restored file count."""
    source = Path(source)
    if not source.is_file() or not zipfile.is_zipfile(source):
        raise ValueError("The selected file is not a valid Planning backup.")
    APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    restored = 0
    with tempfile.TemporaryDirectory(prefix="planing-restore-") as temp:
        temp_root = Path(temp)
        with zipfile.ZipFile(source) as archive:
            for member in archive.infolist():
                name = Path(member.filename)
                if name.is_absolute() or ".." in name.parts or name.name == "backup_info.txt":
                    continue
                archive.extract(member, temp_root)
        for path in temp_root.rglob("*"):
            if not path.is_file():
                continue
            target = APP_DATA_DIR / path.relative_to(temp_root)
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, target)
            restored += 1
    return restored
