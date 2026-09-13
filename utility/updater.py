"""GitHub Releases version checking and installer update helpers."""

from __future__ import annotations

from dataclasses import dataclass
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from typing import Callable
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
import zipfile


DEFAULT_APP_VERSION = "1.0.0"
GITHUB_REPOSITORY = "bibo-crypto/Planing"
RELEASES_API_URL = f"https://api.github.com/repos/{GITHUB_REPOSITORY}/releases/latest"
RELEASES_PAGE_URL = f"https://github.com/{GITHUB_REPOSITORY}/releases"


def version_file_path() -> Path:
    """Return the installed version file, supporting PyInstaller's _internal layout."""
    if getattr(sys, "frozen", False):
        install_dir = Path(sys.executable).resolve().parent
        for candidate in (install_dir / "version.txt", install_dir / "_internal" / "version.txt"):
            if candidate.is_file():
                return candidate
        return install_dir / "version.txt"
    return Path(__file__).resolve().parent.parent / "version.txt"


def get_installed_version() -> str:
    try:
        value = version_file_path().read_text(encoding="utf-8").strip()
        if value:
            return value.lstrip("vV")
    except OSError:
        pass
    return DEFAULT_APP_VERSION


APP_VERSION = get_installed_version()


@dataclass(frozen=True)
class ReleaseInfo:
    version: str
    tag_name: str
    title: str
    notes: str
    page_url: str
    installer_url: str = ""
    installer_name: str = ""


def _version_tuple(value: str) -> tuple[int, ...]:
    """Parse v1.2.3-style versions; non-numeric suffixes are ignored."""
    clean = str(value or "").strip().lstrip("vV")
    parts: list[int] = []
    for part in clean.split("."):
        digits = "".join(ch for ch in part if ch.isdigit())
        if not digits:
            break
        parts.append(int(digits))
    return tuple(parts or [0])


def is_newer_version(candidate: str, current: str | None = None) -> bool:
    current = current or get_installed_version()
    return _version_tuple(candidate) > _version_tuple(current)


def _choose_installer(assets: list[dict]) -> tuple[str, str]:
    candidates = []
    for asset in assets:
        name = str(asset.get("name", ""))
        url = str(asset.get("browser_download_url", ""))
        if url and Path(name).suffix.lower() == ".zip":
            candidates.append((name, url))
    # Prefer the Inno Setup installer when a release contains other EXEs.
    candidates.sort(key=lambda item: ("setup" not in item[0].casefold(), item[0].casefold()))
    return candidates[0] if candidates else ("", "")


def fetch_latest_release(timeout: float = 6.0) -> ReleaseInfo | None:
    """Fetch the latest published GitHub release, returning None on failure."""
    request = Request(
        RELEASES_API_URL,
        headers={"Accept": "application/vnd.github+json", "User-Agent": "Planing-Updater"},
    )
    try:
        with urlopen(request, timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except (HTTPError, URLError, TimeoutError, OSError, ValueError):
        return None

    tag_name = str(payload.get("tag_name", "")).strip()
    version = tag_name.lstrip("vV")
    if not tag_name or not is_newer_version(version, get_installed_version()):
        return None
    installer_name, installer_url = _choose_installer(payload.get("assets") or [])
    return ReleaseInfo(
        version=version,
        tag_name=tag_name,
        title=str(payload.get("name", "") or tag_name),
        notes=str(payload.get("body", "") or "").strip(),
        page_url=str(payload.get("html_url", "") or RELEASES_PAGE_URL),
        installer_url=installer_url,
        installer_name=installer_name,
    )


def check_for_updates_async(callback: Callable[[ReleaseInfo | None], None], logger=None) -> None:
    """Check GitHub in a daemon thread and invoke *callback* when finished."""
    import threading

    def worker() -> None:
        try:
            callback(fetch_latest_release())
        except Exception:  # noqa: BLE001
            if logger:
                logger.exception("Update check failed")

    threading.Thread(target=worker, name="planing-update-check", daemon=True).start()


def download_installer(release: ReleaseInfo, progress: Callable[[int], None] | None = None) -> Path:
    """Download the release ZIP to a temporary file and return its path."""
    if not release.installer_url:
        raise RuntimeError("This release has no Windows installer attached yet.")
    suffix = ".zip"
    fd, raw_path = tempfile.mkstemp(prefix="Planing_Update_", suffix=suffix)
    os.close(fd)
    destination = Path(raw_path)
    request = Request(release.installer_url, headers={"User-Agent": "Planing-Updater"})
    try:
        with urlopen(request, timeout=30) as response, destination.open("wb") as output:
            total = int(response.headers.get("Content-Length", "0") or 0)
            downloaded = 0
            while True:
                chunk = response.read(1024 * 256)
                if not chunk:
                    break
                output.write(chunk)
                downloaded += len(chunk)
                if progress and total:
                    progress(int(downloaded * 100 / total))
    except Exception:
        destination.unlink(missing_ok=True)
        raise
    return destination


def _safe_extract(zip_path: Path, destination: Path) -> None:
    """Extract a ZIP without allowing members to escape the temp folder."""
    destination_resolved = destination.resolve()
    with zipfile.ZipFile(zip_path) as archive:
        for member in archive.infolist():
            target = (destination / member.filename).resolve()
            if target != destination_resolved and destination_resolved not in target.parents:
                raise RuntimeError("The update ZIP contains an unsafe file path.")
        archive.extractall(destination)


def _powershell_quote(value: Path | str) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def install_update(zip_path: Path, install_dir: Path, new_version: str) -> None:
    """Stage a ZIP update and start a hidden process that replaces files later.

    The running EXE cannot replace itself. The generated PowerShell helper
    waits for Planing to exit, copies the new application files over the
    existing installation, then starts the new EXE. User settings/data live
    in AppData and are never touched.
    """
    if not zip_path.is_file():
        raise FileNotFoundError(zip_path)
    if not getattr(sys, "frozen", False):
        raise RuntimeError("Automatic updates are available only in the installed Windows version.")

    stage_root = Path(tempfile.mkdtemp(prefix="Planing_Update_"))
    try:
        _safe_extract(zip_path, stage_root)
        executables = sorted(stage_root.rglob("Planing.exe"), key=lambda path: len(path.parts))
        if not executables:
            raise RuntimeError("The update ZIP must contain Planing.exe.")
        # Accept all common ZIP layouts: Planing.exe at the root, inside a
        # Planing folder, or inside an archived dist/Planing folder.
        source_dir = executables[0].parent

        script_path = stage_root.parent / f"Planing_Update_{os.getpid()}.ps1"
        exe_path = install_dir / "Planing.exe"
        script = f"""$ErrorActionPreference = 'Stop'
$log = Join-Path $env:TEMP 'Planing_Update.log'
Start-Sleep -Seconds 3
$source = {_powershell_quote(source_dir)}
$target = {_powershell_quote(install_dir)}
try {{
    for ($attempt = 1; $attempt -le 20; $attempt++) {{
        try {{
            Get-ChildItem -LiteralPath $source -Force | ForEach-Object {{
                Copy-Item -LiteralPath $_.FullName -Destination $target -Recurse -Force -ErrorAction Stop
            }}
            if (Test-Path -LiteralPath {_powershell_quote(exe_path)}) {{
                Set-Content -LiteralPath (Join-Path $target 'version.txt') -Value {_powershell_quote(new_version)} -Encoding UTF8
                Start-Process -FilePath {_powershell_quote(exe_path)} -WorkingDirectory {_powershell_quote(install_dir)}
                break
            }}
        }} catch {{
            $_ | Out-File -FilePath $log -Append
            Start-Sleep -Seconds 1
        }}
    }}
    }} catch {{ $_ | Out-File -FilePath $log -Append }}
Remove-Item -LiteralPath {_powershell_quote(zip_path)} -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath {_powershell_quote(stage_root)} -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
"""
        script_path.write_text(script, encoding="utf-8")
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", str(script_path)],
            creationflags=creation_flags,
            close_fds=True,
        )
    except Exception:
        shutil.rmtree(stage_root, ignore_errors=True)
        zip_path.unlink(missing_ok=True)
        raise
