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


DEFAULT_APP_VERSION = "1.0.6"
GITHUB_REPOSITORY = "bibo-crypto/Planing"
RELEASES_API_URL = f"https://api.github.com/repos/{GITHUB_REPOSITORY}/releases/latest"
RELEASES_PAGE_URL = f"https://github.com/{GITHUB_REPOSITORY}/releases"


def version_file_path() -> Path:
    """Return the installed version file, supporting PyInstaller's _internal layout."""
    if getattr(sys, "frozen", False):
        install_dir = Path(sys.executable).resolve().parent
        # The external file is authoritative: the updater can replace it
        # without rebuilding the PyInstaller internal archive.
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


def write_version_file(version: str, target: Path | None = None) -> Path:
    """Atomically write the installed version and return the written path."""
    path = target or version_file_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    temporary.write_text(str(version).strip().lstrip("vV") + "\n", encoding="utf-8")
    temporary.replace(path)
    written = path.read_text(encoding="utf-8").strip().lstrip("vV")
    if written != str(version).strip().lstrip("vV"):
        raise RuntimeError(f"The version file could not be verified at {path}.")
    return path


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
        # A network hiccup mid-download can end the response early without
        # urlopen/read ever raising -- silently continuing with a truncated
        # or corrupt ZIP produced exactly this symptom in the wild: the
        # install script would extract *something*, but Windows refused to
        # run it ("not a valid application for this OS platform"), and that
        # only ever surfaced deep in the PowerShell log, long after this
        # function returned successfully. Catch it right here instead.
        if total and downloaded != total:
            raise RuntimeError(
                f"Download incomplete: got {downloaded} of {total} bytes. Check your internet connection and try again."
            )
        try:
            with zipfile.ZipFile(destination) as archive:
                bad_file = archive.testzip()
        except zipfile.BadZipFile as exc:
            raise RuntimeError("The downloaded update file is corrupt. Please try again.") from exc
        if bad_file:
            raise RuntimeError(f"The downloaded update file is corrupt (bad entry: {bad_file}). Please try again.")
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


def _needs_elevation(install_dir: Path) -> bool:
    """True if the current (non-admin) process cannot write into install_dir
    -- e.g. it's under C:\\Program Files, which every normal user account
    (including the one that owns/runs this installation) needs
    Administrator rights to modify. Probed with a real write, not a guess
    from the path alone, so a per-user install location (%LOCALAPPDATA%)
    never triggers an unnecessary UAC prompt."""
    probe = install_dir / f".planing_write_test_{os.getpid()}.tmp"
    try:
        probe.write_text("x", encoding="utf-8")
        probe.unlink(missing_ok=True)
        return False
    except OSError:
        return True


def _run_update_script(script_path: Path, creation_flags: int, elevated: bool) -> None:
    if not elevated:
        subprocess.Popen(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", str(script_path)],
            creationflags=creation_flags,
            close_fds=True,
        )
        return
    # install_dir (typically C:\Program Files\Planing) needs Administrator
    # rights to write to, same as any other app installed there -- this is
    # not optional, so request elevation via the UAC prompt (the "runas"
    # verb) rather than letting every Copy-Item in the script fail with
    # Access Denied, including the rollback's own copy-back, which used to
    # leave nothing running at all (see install_update's docstring).
    import ctypes
    args = f'-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{script_path}"'
    result = ctypes.windll.shell32.ShellExecuteW(None, "runas", "powershell.exe", args, None, 0)
    if result <= 32:  # ShellExecuteW's own convention: <=32 means it failed to even launch
        raise RuntimeError(
            "This update needs Administrator approval to install (Planing is installed under "
            "Program Files). Please approve the Windows prompt, or ask an admin to run the update."
        )


def install_update(zip_path: Path, install_dir: Path, new_version: str) -> None:
    """Stage a ZIP update and start a hidden process that replaces files later.

    The running EXE cannot replace itself. The generated PowerShell helper
    waits for Planing to exit, backs up the current installation, copies
    the new application files over it, starts the new EXE, and verifies it
    actually stays running -- if the new build crashes immediately on
    launch, or the file copy never succeeds, it automatically restores the
    backup and relaunches the previous, known-working version instead of
    leaving the install broken or, worse, leaving nothing running at all.
    User settings/data live in AppData and are never touched by either the
    update or the rollback.

    install_dir commonly needs Administrator rights (Program Files) that
    this process does not itself have, so the helper script is launched
    elevated (UAC prompt) whenever a real write probe shows install_dir
    isn't writable as-is -- otherwise every step below fails with Access
    Denied, including the rollback's own copy-back, silently.
    """
    if not zip_path.is_file():
        raise FileNotFoundError(zip_path)
    if not getattr(sys, "frozen", False):
        raise RuntimeError("Automatic updates are available only in the installed Windows version.")

    elevated = _needs_elevation(install_dir)
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
        # A sibling of install_dir, not inside it, so clearing/restoring
        # install_dir during a rollback never touches the backup itself.
        # Overwritten fresh on every update -- always exactly one, most
        # recent known-good version to fall back to.
        backup_dir = install_dir.parent / "Planing_Backup"
        script = f"""$ErrorActionPreference = 'Stop'
$log = Join-Path $env:TEMP 'Planing_Update.log'
# Give the old process time to fully exit and release its file locks, then
# explicitly wait until no old Planing process remains before copying. The
# fixed delay alone was not sufficient on slower machines/antivirus scans.
Start-Sleep -Seconds 5
for ($check = 1; $check -le 30; $check++) {{
    if (-not (Get-Process -Name 'Planing' -ErrorAction SilentlyContinue)) {{ break }}
    Start-Sleep -Seconds 2
}}
$source = {_powershell_quote(source_dir)}
$target = {_powershell_quote(install_dir)}
$backup = {_powershell_quote(backup_dir)}
$exePath = {_powershell_quote(exe_path)}
function Restore-Backup {{
    # Never let a rollback failure propagate out under $ErrorActionPreference
    # = 'Stop' -- that would skip the Start-Process relaunch right after it
    # and leave nothing running at all. Best-effort is still far better than
    # a silent, total failure here.
    try {{
        Get-ChildItem -LiteralPath $target -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        Get-ChildItem -LiteralPath $backup -Force | ForEach-Object {{
            Copy-Item -LiteralPath $_.FullName -Destination $target -Recurse -Force -ErrorAction SilentlyContinue
        }}
    }} catch {{ $_ | Out-File -FilePath $log -Append }}
}}
try {{
    if (Test-Path -LiteralPath $backup) {{ Remove-Item -LiteralPath $backup -Recurse -Force -ErrorAction SilentlyContinue }}
    New-Item -ItemType Directory -Path $backup -Force | Out-Null
    Get-ChildItem -LiteralPath $target -Force | ForEach-Object {{
        Copy-Item -LiteralPath $_.FullName -Destination $backup -Recurse -Force
    }}

    $copied = $false
    for ($attempt = 1; $attempt -le 20; $attempt++) {{
        try {{
            Get-ChildItem -LiteralPath $source -Force | ForEach-Object {{
                Copy-Item -LiteralPath $_.FullName -Destination $target -Recurse -Force -ErrorAction Stop
            }}
            if (Test-Path -LiteralPath $exePath) {{
                $copied = $true
                break
            }}
        }} catch {{
            $_ | Out-File -FilePath $log -Append
            Start-Sleep -Seconds 2
        }}
    }}

    if ($copied) {{
        $versionPath = Join-Path $target 'version.txt'
        $versionValue = {_powershell_quote(new_version)}
        $versionTemp = Join-Path $target ('.version.' + $PID + '.tmp')
        Set-Content -LiteralPath $versionTemp -Value $versionValue -Encoding UTF8
        Move-Item -LiteralPath $versionTemp -Destination $versionPath -Force
        $writtenVersion = (Get-Content -LiteralPath $versionPath -Raw).Trim().TrimStart('v','V')
        if ($writtenVersion -ne $versionValue.Trim().TrimStart('v','V')) {{ throw "version.txt verification failed: '$writtenVersion'" }}
        Start-Process -FilePath $exePath -WorkingDirectory $target
        # Antivirus commonly intercepts a freshly-written, unsigned exe on
        # its first launch -- it can run its own scan/sandbox pass, kill
        # that first process, and start a *different* one afterwards, on
        # top of an already slower cold start. Watching one specific PID
        # (Start-Process -PassThru) would misread that AV cycle as a crash,
        # and a short fixed wait isn't enough either -- this has been
        # observed taking close to a minute in practice. So instead: poll
        # by process name (survives the PID changing under AV) for up to 3
        # minutes, and succeed the moment Planing.exe is seen running even
        # once -- a genuine crash never shows up running at all, no matter
        # how long or how many times checked.
        $sawRunning = $false
        for ($check = 1; $check -le 36; $check++) {{
            Start-Sleep -Seconds 5
            if (Get-Process -Name 'Planing' -ErrorAction SilentlyContinue) {{
                $sawRunning = $true
                break
            }}
        }}
        if (-not $sawRunning) {{
            "New version never started running within 3 minutes -- rolling back to the previous version" | Out-File -FilePath $log -Append
            Restore-Backup
            Start-Process -FilePath $exePath -WorkingDirectory $target -ErrorAction SilentlyContinue
        }}
    }} else {{
        "Update copy never succeeded after 20 attempts -- rolling back to the previous version" | Out-File -FilePath $log -Append
        Restore-Backup
        Start-Process -FilePath $exePath -WorkingDirectory $target -ErrorAction SilentlyContinue
    }}
    }} catch {{
        $_ | Out-File -FilePath $log -Append
        # Whatever failed above, always try to leave *something* running --
        # closing Planing to update and then leaving nothing open at all,
        # not even the old version, is the one outcome worse than a failed
        # update.
        Restore-Backup
        Start-Process -FilePath $exePath -WorkingDirectory $target -ErrorAction SilentlyContinue
    }}
Remove-Item -LiteralPath {_powershell_quote(zip_path)} -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath {_powershell_quote(stage_root)} -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
"""
        script_path.write_text(script, encoding="utf-8")
        # CREATE_NO_WINDOW alone only hides the console -- it doesn't
        # detach the process from anything. If Planing.exe is running
        # inside a Windows Job Object (common under some launchers, remote
        # sessions, or security software), closing the parent can silently
        # kill this helper along with it before it ever gets to run,
        # leaving the update never applied and nothing relaunched at all.
        # CREATE_BREAKAWAY_FROM_JOB + CREATE_NEW_PROCESS_GROUP make sure
        # this process survives the parent's exit regardless. Irrelevant
        # (and unavailable) for the elevated/ShellExecuteW path -- a
        # UAC-elevated process is already its own top-level session.
        creation_flags = (
            getattr(subprocess, "CREATE_NO_WINDOW", 0)
            | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            | getattr(subprocess, "CREATE_BREAKAWAY_FROM_JOB", 0)
        )
        _run_update_script(script_path, creation_flags, elevated)
    except Exception:
        shutil.rmtree(stage_root, ignore_errors=True)
        zip_path.unlink(missing_ok=True)
        raise
