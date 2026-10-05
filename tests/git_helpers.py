"""Small git helpers shared by the git/publish tests."""
import subprocess
from pathlib import Path


def git(cwd, *args, check=True):
    proc = subprocess.run(["git", *args], cwd=str(cwd), capture_output=True, text=True)
    if check and proc.returncode != 0:
        raise AssertionError(f"git {' '.join(args)} failed: {proc.stderr}")
    return proc.stdout.strip()


def write(root, relative, text="x\n"):
    path = Path(root) / relative
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


def remote_log(repo):
    remote = git(repo, "remote", "get-url", "origin")
    return git(remote, "log", "--format=%s", "main").splitlines()
