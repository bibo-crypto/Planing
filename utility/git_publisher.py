"""Commit and push the project's changes to its Git remote.

Used by ``publish.py`` / ``publish.bat``. Everything here is plain Python
around the ``git`` command line (standard library only, no Tk, no other
project module), so it can be tested against a real throw-away repository.

Design rules, all aimed at making a one-click push hard to get wrong:

* ``.gitignore`` is always respected. Changes come from ``git status`` (which
  never lists ignored files) and are staged with a plain ``git add`` -- never
  ``--force``.
* Only the files the user selected are committed (``git commit`` with an
  explicit pathspec), even if other things happen to be staged already.
* Never force-pushes, never rewrites history, never pulls. If the remote has
  newer commits the push is refused and the reason is explained.
* Files that look like secrets or are too large for GitHub are flagged and
  unchecked by default.
* Credentials are never stored or typed here: git's own credential helper
  (Git Credential Manager on Windows) is used, and any user/token embedded in
  a remote URL is masked before it is shown.
"""
from __future__ import annotations

import logging
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Callable

logger = logging.getLogger("planing.git")

COMMIT_TYPES = ("feat", "fix", "refactor", "perf", "docs", "test", "build", "ci", "chore", "update")

# GitHub warns above 50 MB and rejects single files above 100 MB.
LARGE_FILE_BYTES = 50 * 1024 * 1024

_NETWORK_TIMEOUT = 180
_LOCAL_TIMEOUT = 60

_SECRET_NAME = re.compile(
    r"(^|/)("
    r"\.env(\..+)?|id_rsa.*|id_ed25519.*|.+\.pem|.+\.pfx|.+\.p12|.+\.key"
    r"|credentials?(\..+)?|.*secrets?.*\.(json|ya?ml|txt|ini)|.*tokens?.*\.(json|txt)"
    r")$",
    re.IGNORECASE,
)

_BUILD_FILES = {
    "build.bat", "main.spec", "installer.iss", "requirements.txt",
    "runtime_hook.py", "sync_venv_packages.py",
}


class GitError(Exception):
    """A git operation failed; the message is written for the end user."""


class PushFailedError(GitError):
    """The commit was created locally but pushing it failed."""

    def __init__(self, message: str, commit: str = ""):
        super().__init__(message)
        self.commit = commit


@dataclass
class Change:
    path: str                 # forward-slash path relative to the repo root
    status: str               # "A" new, "M" modified, "D" deleted
    added: int | None = None  # lines added (None: binary / unknown)
    deleted: int | None = None
    size: int = 0             # bytes on disk (0 for deleted files)
    risk: str = ""            # non-empty: why it should not be committed blindly


@dataclass
class RepoInfo:
    root: Path
    branch: str
    remote: str = ""
    remote_url: str = ""      # credentials masked
    upstream: str = ""        # e.g. "origin/main" ('' if the branch has none yet)
    ahead: int = 0            # local commits not on the remote
    behind: int = 0           # remote commits not here
    has_commits: bool = True
    fetch_error: str = ""


@dataclass
class CommitMessage:
    type: str
    scope: str
    summary: str
    body: str

    @property
    def subject(self) -> str:
        head = f"{self.type}({self.scope})" if self.scope else self.type
        return f"{head}: {self.summary}".strip()

    def full(self) -> str:
        return f"{self.subject}\n\n{self.body}".strip() + "\n" if self.body else self.subject + "\n"


@dataclass
class PublishResult:
    commit: str = ""          # short hash of the new commit ('' for push-only)
    pushed: bool = False
    branch: str = ""
    remote_url: str = ""
    output: str = ""


# ---------------------------------------------------------------------------
# Running git
# ---------------------------------------------------------------------------

def find_git() -> str | None:
    """Path of the git executable, or None if git isn't installed."""
    found = shutil.which("git")
    if found:
        return found
    if os.name == "nt":
        for env_name in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
            base = os.environ.get(env_name)
            if not base:
                continue
            for relative in ("Git/cmd/git.exe", "Programs/Git/cmd/git.exe"):
                candidate = Path(base) / relative
                if candidate.is_file():
                    return str(candidate)
    return None


def explain_git_error(text: str) -> str:
    """Turn git's stderr into something an end user can act on."""
    raw = (text or "").strip()
    low = raw.lower()
    if "non-fast-forward" in low or "fetch first" in low or "tip of your current branch is behind" in low \
            or ("rejected" in low and "remote contains work" in low):
        return ("The remote has commits that you don't have yet, so GitHub refused the push.\n"
                "Update first (git pull --rebase), then publish again.")
    if "authentication failed" in low or "could not read username" in low or "terminal prompts disabled" in low \
            or "permission denied" in low or "invalid username or password" in low \
            or "returned error: 403" in low or "repository not found" in low:
        return ("GitHub did not accept the credentials (or you have no write access to the repository).\n"
                "Sign in once through Git Credential Manager (run `git push` in a terminal), then try again.")
    if "please tell me who you are" in low or "unable to auto-detect email" in low or "empty ident name" in low:
        return ("Git doesn't know who you are. Run once:\n"
                "git config --global user.name \"Your Name\"\n"
                "git config --global user.email \"you@example.com\"")
    if "dubious ownership" in low:
        return ("Git refuses to use this folder because of its ownership. Run once:\n"
                "git config --global --add safe.directory <this folder>")
    if "could not resolve host" in low or "unable to access" in low or "network is unreachable" in low \
            or "connection timed out" in low or "failed to connect" in low:
        return "Couldn't reach GitHub. Check the internet connection and try again."
    if "gh001" in low or "file size limit" in low or "large files detected" in low:
        return "GitHub rejected a file because it is too large (limit 100 MB). Remove it from the commit and try again."
    if "nothing to commit" in low or "no changes added to commit" in low or "nothing added to commit" in low:
        return "There is nothing to commit for the selected files."
    return raw or "Git failed without giving a reason."


def _run(root: Path | str, args: list[str], *, timeout: int = _LOCAL_TIMEOUT) -> subprocess.CompletedProcess:
    git = find_git()
    if not git:
        raise GitError("Git is not installed on this computer (https://git-scm.com/downloads).")
    env = dict(os.environ)
    env["GIT_TERMINAL_PROMPT"] = "0"   # fail instead of waiting on a prompt nobody can see
    env["LC_ALL"] = "C"                # English messages, so explain_git_error() can recognise them
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
    try:
        return subprocess.run(
            [git, "-c", "core.quotepath=false", *args],
            cwd=str(root), capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=timeout, env=env, creationflags=flags,
        )
    except subprocess.TimeoutExpired as exc:
        raise GitError(f"`git {args[0]}` took longer than {timeout} seconds and was stopped.") from exc
    except OSError as exc:
        raise GitError(f"Couldn't run git: {exc}") from exc


def _git(root: Path | str, *args: str, timeout: int = _LOCAL_TIMEOUT) -> str:
    """Run git and return stdout; raise GitError (with a friendly message) on failure."""
    proc = _run(root, list(args), timeout=timeout)
    if proc.returncode != 0:
        raise GitError(explain_git_error(proc.stderr or proc.stdout))
    return proc.stdout


def mask_url(url: str) -> str:
    """Hide any user/password/token embedded in a remote URL."""
    return re.sub(r"(?<=://)[^/@\s]+@", "", url or "")


# ---------------------------------------------------------------------------
# Repository discovery and state
# ---------------------------------------------------------------------------

def resolve_repo(path: Path | str) -> Path:
    """Top-level folder of the repository containing *path*."""
    folder = Path(path)
    if not folder.is_dir():
        raise GitError(f"Folder not found: {folder}")
    proc = _run(folder, ["rev-parse", "--show-toplevel"])
    if proc.returncode != 0:
        raise GitError(explain_git_error(proc.stderr) if "dubious" in (proc.stderr or "").lower()
                       else "This folder is not inside a Git repository.")
    return Path(proc.stdout.strip())


def _has_commits(root: Path) -> bool:
    return _run(root, ["rev-parse", "--verify", "-q", "HEAD"]).returncode == 0


def repo_info(root: Path, *, fetch: bool = True) -> RepoInfo:
    """Branch, remote and how far ahead/behind the branch is.

    With *fetch* the remote is contacted first so "behind" is current; a failed
    fetch (offline) is reported in ``fetch_error`` rather than raised.
    """
    proc = _run(root, ["symbolic-ref", "--short", "-q", "HEAD"])
    branch = proc.stdout.strip()
    if proc.returncode != 0 or not branch:
        raise GitError("The repository is on a detached HEAD (no branch is checked out). "
                       "Switch to a branch first.")
    info = RepoInfo(root=root, branch=branch, has_commits=_has_commits(root))

    remote = _run(root, ["config", "--get", f"branch.{branch}.remote"]).stdout.strip()
    if not remote:
        remotes = _run(root, ["remote"]).stdout.split()
        remote = "origin" if "origin" in remotes else (remotes[0] if remotes else "")
    info.remote = remote
    if remote:
        info.remote_url = mask_url(_run(root, ["remote", "get-url", remote]).stdout.strip())
        if fetch:
            fetched = _run(root, ["fetch", "--quiet", remote], timeout=_NETWORK_TIMEOUT)
            if fetched.returncode != 0:
                info.fetch_error = explain_git_error(fetched.stderr)

    upstream = _run(root, ["rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}"])
    if upstream.returncode == 0:
        info.upstream = upstream.stdout.strip()
        counts = _run(root, ["rev-list", "--left-right", "--count", "HEAD...@{u}"]).stdout.split()
        if len(counts) == 2:
            info.ahead, info.behind = int(counts[0]), int(counts[1])
    return info


def _risk_of(path: str, size: int) -> str:
    if size > LARGE_FILE_BYTES:
        return f"larger than {LARGE_FILE_BYTES // (1024 * 1024)} MB (GitHub rejects files over 100 MB)"
    if _SECRET_NAME.search(path):
        return "looks like a secret or credentials file"
    return ""


def _count_lines(file_path: Path) -> int | None:
    try:
        if file_path.stat().st_size > 2 * 1024 * 1024:
            return None
        data = file_path.read_bytes()
    except OSError:
        return None
    if b"\0" in data:
        return None  # binary
    return data.count(b"\n") + (1 if data and not data.endswith(b"\n") else 0)


def list_changes(root: Path) -> list[Change]:
    """Every uncommitted change that git would track (ignored files never appear)."""
    out = _git(root, "status", "--porcelain=v1", "-z", "--untracked-files=all", "--no-renames")
    entries: list[tuple[str, str]] = []
    for item in out.split("\0"):
        if len(item) < 4:
            continue
        xy, path = item[:2], item[3:]
        if xy in {"DD", "AU", "UD", "UA", "DU", "AA", "UU"}:
            raise GitError("There are unresolved merge conflicts. Resolve them first.")
        entries.append((xy, path))

    stats: dict[str, tuple[int | None, int | None]] = {}
    if entries and _has_commits(root):
        for line in _git(root, "diff", "HEAD", "--numstat", "--no-renames").splitlines():
            parts = line.split("\t", 2)
            if len(parts) == 3:
                added, deleted, path = parts
                stats[path] = (int(added) if added.isdigit() else None,
                               int(deleted) if deleted.isdigit() else None)

    changes: list[Change] = []
    for xy, path in sorted(entries, key=lambda e: e[1].lower()):
        if xy == "??" or "A" in xy:
            status = "A"
        elif "D" in xy:
            status = "D"
        else:
            status = "M"
        file_path = root / path
        size = file_path.stat().st_size if status != "D" and file_path.is_file() else 0
        added, deleted = stats.get(path, (None, None))
        if status == "A" and path not in stats:
            added, deleted = _count_lines(file_path), 0
        changes.append(Change(path=path, status=status, added=added, deleted=deleted,
                              size=size, risk=_risk_of(path, size)))
    return changes


# ---------------------------------------------------------------------------
# Commit message
# ---------------------------------------------------------------------------

_AREA_SUFFIXES = ("_tab", "_exporter", "_refresh", "_sources", "_window", "_actions", "_workflow",
                  "_db", "_cache", "_logic")
_AREA_SPECIAL = {
    "gui": "gui", "modern_widgets": "gui", "main": "app", "utils": "utils", "updater": "updater",
    "notifications": "notifications", "disk_cache": "disk-cache", "version": "release",
}


def area_of(path: str) -> str:
    """Short feature name for a file, used as the commit scope."""
    p = path.replace("\\", "/")
    name = PurePosixPath(p).name
    if p.startswith(".github/"):
        return "ci"
    if p.startswith("tests/"):
        return "tests"
    if name in _BUILD_FILES:
        return "build"
    if p.startswith("data/"):
        return "data"
    if p.lower().endswith((".md", ".rst")):
        return "docs"
    stem = PurePosixPath(p).stem
    if stem in _AREA_SPECIAL:
        return _AREA_SPECIAL[stem]
    for suffix in _AREA_SUFFIXES:
        if stem.endswith(suffix) and len(stem) > len(suffix):
            stem = stem[: -len(suffix)]
            break
    return stem.replace("_", "-")


def _stat_text(change: Change) -> str:
    parts = []
    if change.added:
        parts.append(f"+{change.added}")
    if change.deleted:
        parts.append(f"-{change.deleted}")
    return f" ({'/'.join(parts)})" if parts else ""


def build_commit_message(changes: list[Change], version: str | None = None) -> CommitMessage:
    """Describe *changes* in a conventional-commit style message.

    The type and wording come from which files changed and how (new / modified /
    deleted, lines added and removed) -- git can't know the intent, so the
    dialog lets the user adjust the type and text before committing.
    """
    if not changes:
        return CommitMessage("update", "", "no changes", "")

    areas = {c.path: area_of(c.path) for c in changes}
    paths = [c.path for c in changes]
    unique_areas = set(areas.values())

    weight: dict[str, int] = {}
    for change in changes:
        weight[areas[change.path]] = weight.get(areas[change.path], 0) + max(
            1, (change.added or 0) + (change.deleted or 0))
    ranked = [a for a, _ in sorted(weight.items(), key=lambda kv: (-kv[1], kv[0]))]
    scoped = [a for a in ranked if a != "tests"] or ranked   # tests ride along with the code they cover
    scope = ", ".join(scoped[:3])

    added = [c for c in changes if c.status == "A"]
    modified = [c for c in changes if c.status == "M"]
    deleted = [c for c in changes if c.status == "D"]
    total_added = sum(c.added or 0 for c in changes)
    total_deleted = sum(c.deleted or 0 for c in changes)

    body_lines = [f"- {c.status} {c.path}{_stat_text(c)}" for c in changes[:40]]
    if len(changes) > 40:
        body_lines.append(f"- ... and {len(changes) - 40} more files")

    if version and "version.txt" in paths:
        others = len(changes) - 1
        body = "\n".join(body_lines) if others else ""
        return CommitMessage("chore", "release", f"bump version to {version}", body)

    if unique_areas == {"tests"}:
        ctype = "test"
    elif unique_areas == {"docs"}:
        ctype = "docs"
    elif unique_areas == {"ci"}:
        ctype = "ci"
    elif unique_areas <= {"build"}:
        ctype = "build"
    elif any(c.path.endswith(".py") and areas[c.path] != "tests" for c in added):
        ctype = "feat"
    else:
        ctype = "update"

    if len(changes) == 1:
        change = changes[0]
        verb = {"A": "add", "M": "update", "D": "remove"}[change.status]
        summary = f"{verb} {PurePosixPath(change.path).name}"
    elif ctype == "feat" and len(added) <= 2 and not deleted:
        names = " and ".join(PurePosixPath(c.path).name for c in added if c.path.endswith(".py")) \
            or " and ".join(PurePosixPath(c.path).name for c in added)
        extra = len(changes) - len(added)
        summary = f"add {names}" + (f" and update {extra} file{'s' if extra != 1 else ''}" if extra else "")
    else:
        counts = []
        if added:
            counts.append(f"{len(added)} added")
        if modified:
            counts.append(f"{len(modified)} modified")
        if deleted:
            counts.append(f"{len(deleted)} deleted")
        summary = ", ".join(counts)
        if total_added or total_deleted:
            summary += f" (+{total_added}/-{total_deleted})"

    return CommitMessage(ctype, scope, summary, "\n".join(body_lines))


# ---------------------------------------------------------------------------
# Commit and push
# ---------------------------------------------------------------------------

def _write_temp(data: bytes) -> str:
    handle, name = tempfile.mkstemp(prefix="planing-git-", suffix=".tmp")
    with os.fdopen(handle, "wb") as fh:
        fh.write(data)
    return name


def push(root: Path, progress: Callable[[str], None] | None = None) -> RepoInfo:
    """Push the current branch (setting its upstream the first time)."""
    info = repo_info(root, fetch=False)
    if not info.remote:
        raise GitError("This repository has no remote (nowhere to push to). "
                       "Add one with: git remote add origin <url>")
    if progress:
        progress(f"Pushing {info.branch} to {info.remote}...")
    args = ["push"] if info.upstream else ["push", "-u", info.remote, info.branch]
    proc = _run(root, args, timeout=_NETWORK_TIMEOUT)
    if proc.returncode != 0:
        raise GitError(explain_git_error(proc.stderr or proc.stdout))
    return info


def publish(root: Path, paths: list[str], message: str, *, do_push: bool = True,
            progress: Callable[[str], None] | None = None) -> PublishResult:
    """Commit exactly *paths* with *message*, then (optionally) push.

    With no *paths* and *do_push* this only pushes commits that already exist
    locally. A failed push leaves the commit in place and raises
    PushFailedError so nothing is lost.
    """
    say = progress or (lambda _text: None)
    result = PublishResult()
    commit_made = False

    if paths:
        message = (message or "").strip()
        if not message:
            raise GitError("The commit message is empty.")
        has_head = _has_commits(root)
        pathspec = _write_temp(b"\0".join(p.encode("utf-8") for p in paths) + b"\0")
        message_file = _write_temp(message.encode("utf-8") + b"\n")
        try:
            say(f"Staging {len(paths)} file(s)...")
            # Plain `git add`: .gitignore is honoured, nothing is ever forced in.
            _git(root, "add", "-A", "--pathspec-from-file=" + pathspec, "--pathspec-file-nul")
            say("Creating the commit...")
            commit_args = ["commit", "-F", message_file]
            if has_head:   # an explicit pathspec limits the commit to the selected files
                commit_args += ["--pathspec-from-file=" + pathspec, "--pathspec-file-nul"]
            _git(root, *commit_args)
            commit_made = True
        finally:
            for temp in (pathspec, message_file):
                try:
                    os.unlink(temp)
                except OSError:
                    pass
        result.commit = _git(root, "rev-parse", "--short", "HEAD").strip()
        logger.info("Git: committed %s (%d file(s)): %s", result.commit, len(paths), message.splitlines()[0])

    if do_push:
        try:
            info = push(root, progress)
        except GitError as exc:
            if commit_made:
                raise PushFailedError(
                    f"The commit {result.commit} was created on this computer, but publishing it failed:\n\n{exc}",
                    result.commit,
                ) from exc
            raise
        result.pushed = True
        result.branch = info.branch
        result.remote_url = info.remote_url
        logger.info("Git: pushed %s to %s", info.branch, info.remote_url or info.remote)
    return result
