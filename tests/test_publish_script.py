"""Tests for publish.py (the console tool behind publish.bat), using real git repositories."""
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

import publish
from tests.git_helpers import git, remote_log, write
from utility import git_publisher as gp

pytestmark = pytest.mark.skipif(gp.find_git() is None, reason="git is not installed")


class Console:
    """Scripted answers for input() and a transcript of everything printed."""

    def __init__(self, *answers):
        self.answers = list(answers)
        self.prompts = []
        self.lines = []

    def ask(self, prompt):
        self.prompts.append(prompt)
        if not self.answers:
            raise EOFError
        return self.answers.pop(0)

    def out(self, text=""):
        self.lines.append(text)

    @property
    def text(self):
        return "\n".join(self.lines)


def run(repo_path, *argv, answers=()):
    console = Console(*answers)
    code = publish.main(list(argv), ask=console.ask, out=console.out, root=repo_path)
    return code, console


def head_files(repo_path):
    return git(repo_path, "-c", "core.quotepath=false", "show", "--name-only", "--format=", "HEAD").splitlines()


# ---------------------------------------------------------------- non-interactive

def test_yes_commits_every_file_with_a_generated_message_and_pushes(repo):
    write(repo, "calculate/prezzi.py", "a = 2\n")
    write(repo, "calculate/new_tool.py", "t = 1\n")

    code, console = run(repo, "--yes")

    assert code == 0
    assert remote_log(repo)[0] == "feat(prezzi, new-tool): add new_tool.py and update 1 file"
    assert sorted(head_files(repo)) == ["calculate/new_tool.py", "calculate/prezzi.py"]
    assert "Done. Created commit" in console.text and "pushed main" in console.text
    assert console.prompts == []                       # --yes never asks anything


def test_gitignored_files_are_never_listed_or_committed(repo):
    write(repo, "settings/prefs.json", "{}")
    write(repo, "report.xlsx", "x")
    write(repo, "data/orders.sqlite3", "db")
    write(repo, "calculate/prezzi.py", "a = 2\n")

    code, console = run(repo, "--yes")

    assert code == 0
    assert head_files(repo) == ["calculate/prezzi.py"]
    assert "settings" not in console.text and "report.xlsx" not in console.text and ".sqlite3" not in console.text
    assert set(git(repo, "ls-files").splitlines()) == {".gitignore", "calculate/prezzi.py", "gui/gui.py"}


def test_yes_leaves_out_risky_files_and_says_so(repo):
    write(repo, ".env", "TOKEN=1")
    write(repo, "calculate/prezzi.py", "a = 2\n")

    code, console = run(repo, "--yes")

    assert code == 0
    assert head_files(repo) == ["calculate/prezzi.py"]
    assert "[ ]" in console.text and "WARNING: looks like a secret" in console.text
    assert [c.path for c in gp.list_changes(repo)] == [".env"]      # still waiting, untouched


def test_dry_run_changes_nothing(repo):
    write(repo, "calculate/prezzi.py", "a = 2\n")

    code, console = run(repo, "--dry-run")

    assert code == 0 and "Dry run: would commit 1 file(s) and push." in console.text
    assert git(repo, "log", "-1", "--format=%s") == "initial"
    assert [c.path for c in gp.list_changes(repo)] == ["calculate/prezzi.py"]
    assert console.prompts == []


def test_custom_subject_is_used_and_the_file_list_stays_in_the_body(repo):
    write(repo, "calculate/prezzi.py", "a = 2\nb = 3\n")

    code, _ = run(repo, "--yes", "-m", "fix category ties")

    assert code == 0
    assert git(repo, "log", "-1", "--format=%s") == "fix category ties"
    assert "- M calculate/prezzi.py (+2/-1)" in git(repo, "log", "-1", "--format=%b")


def test_no_push_commits_locally_only(repo):
    write(repo, "calculate/prezzi.py", "a = 2\n")

    code, console = run(repo, "--yes", "--no-push")

    assert code == 0 and "Done. Created commit" in console.text and "and pushed" not in console.text
    assert git(repo, "log", "-1", "--format=%s") != "initial"
    assert remote_log(repo)[0] == "initial"


def test_version_bump_message(repo):
    write(repo, "version.txt", "1.2.0\n")

    code, _ = run(repo, "--yes")

    assert code == 0 and remote_log(repo)[0] == "chore(release): bump version to 1.2.0"


def test_nothing_to_publish(repo):
    code, console = run(repo, "--yes")

    assert code == 0 and "Nothing to publish" in console.text


def test_waiting_commits_are_pushed_when_there_are_no_file_changes(repo):
    write(repo, "x.py"); git(repo, "add", "-A"); git(repo, "commit", "-qm", "waiting")

    code, console = run(repo, "--yes")

    assert code == 0 and "1 commit(s) are waiting to be pushed" in console.text
    assert remote_log(repo)[0] == "waiting"


def test_not_a_repository(tmp_path):
    plain = tmp_path / "plain"
    plain.mkdir()

    code, console = run(plain, "--yes")

    assert code == 1 and "not inside a Git repository" in console.text


def test_github_ahead_means_commit_is_kept_but_push_is_skipped(repo, tmp_path):
    other = tmp_path / "other"
    git(tmp_path, "clone", "-q", git(repo, "remote", "get-url", "origin"), str(other))
    git(other, "config", "user.name", "B"); git(other, "config", "user.email", "b@example.com")
    write(other, "remote_only.py"); git(other, "add", "-A"); git(other, "commit", "-qm", "remote commit"); git(other, "push", "-q")
    write(repo, "calculate/prezzi.py", "a = 2\n")

    code, console = run(repo, "--yes")

    assert code == 1                                   # not fully done
    assert "GitHub has 1 newer commit(s)" in console.text and "git pull --rebase" in console.text
    assert git(repo, "log", "-1", "--format=%s") != "initial"        # the commit was still made
    assert remote_log(repo)[0] == "remote commit"                    # nothing forced onto GitHub


def test_no_remote_commits_locally_and_reports_it(tmp_path):
    solo = tmp_path / "solo"
    git(tmp_path, "init", "-q", "-b", "main", str(solo))
    git(solo, "config", "user.name", "T"); git(solo, "config", "user.email", "t@example.com")
    write(solo, "a.py")

    code, console = run(solo, "--yes")

    assert code == 1 and "no remote" in console.text.lower()
    assert git(solo, "log", "-1", "--format=%s").startswith("feat")


def test_unicode_file_names_are_listed_and_committed(repo):
    write(repo, "docs/ملاحظات.md", "مرحبا\n")

    code, console = run(repo, "--yes")

    assert code == 0 and "docs/ملاحظات.md" in console.text
    assert head_files(repo) == ["docs/ملاحظات.md"]


# ---------------------------------------------------------------- interactive

def test_interactive_enter_accepts_everything(repo):
    write(repo, "calculate/prezzi.py", "a = 2\n")

    code, console = run(repo, answers=["", "", ""])        # files, message, final confirmation

    assert code == 0 and remote_log(repo)[0] == "update(prezzi): update prezzi.py"
    assert len(console.prompts) == 3


def test_interactive_numbers_switch_files_off_and_on(repo):
    write(repo, "calculate/prezzi.py", "a = 2\n")
    write(repo, "gui/gui.py", "b = 2\n")
    write(repo, "utility/other.py", "o = 1\n")

    code, console = run(repo, answers=["1, 3", "", "", ""])    # switch OFF files 1 and 3

    assert code == 0
    assert head_files(repo) == ["gui/gui.py"]
    assert {c.path for c in gp.list_changes(repo)} == {"calculate/prezzi.py", "utility/other.py"}
    assert "[ ]" in console.text and "[x]" in console.text


def test_interactive_custom_subject(repo):
    write(repo, "calculate/prezzi.py", "a = 2\n")

    code, _ = run(repo, answers=["", "fix: my own words", ""])

    assert code == 0 and remote_log(repo)[0] == "fix: my own words"


def test_interactive_quit_changes_nothing(repo):
    write(repo, "calculate/prezzi.py", "a = 2\n")
    for answers in (["q"], ["", "q"], ["", "", "n"]):
        code, console = run(repo, answers=answers)
        assert code == 0 and "Nothing was committed" in console.text
    assert git(repo, "log", "-1", "--format=%s") == "initial"


def test_interactive_risky_file_needs_an_explicit_yes(repo):
    write(repo, ".env", "TOKEN=1")
    write(repo, "calculate/prezzi.py", "a = 2\n")

    code, _ = run(repo, answers=["1", "", "n"])            # switch the .env ON, continue, then refuse the warning
    assert code == 0 and git(repo, "log", "-1", "--format=%s") == "initial"

    code, _ = run(repo, answers=["1", "", "y", "", ""])    # switch on, continue, confirm the warning, message, go
    assert code == 0 and ".env" in head_files(repo)


def test_interactive_with_no_console_input_fails_cleanly(repo):
    write(repo, "calculate/prezzi.py", "a = 2\n")

    code, console = run(repo)                                # input() raises EOFError

    assert code == 1 and "--yes" in console.text
    assert git(repo, "log", "-1", "--format=%s") == "initial"


def test_interactive_ignores_bad_numbers(repo):
    write(repo, "calculate/prezzi.py", "a = 2\n")

    code, console = run(repo, answers=["9 abc", "", "", ""])

    assert code == 0 and "ignored '9'" in console.text and "ignored 'abc'" in console.text


def test_interactive_commit_only_when_github_is_ahead(repo, tmp_path):
    other = tmp_path / "other"
    git(tmp_path, "clone", "-q", git(repo, "remote", "get-url", "origin"), str(other))
    git(other, "config", "user.name", "B"); git(other, "config", "user.email", "b@example.com")
    write(other, "remote_only.py"); git(other, "add", "-A"); git(other, "commit", "-qm", "remote commit"); git(other, "push", "-q")
    write(repo, "calculate/prezzi.py", "a = 2\n")

    code, _ = run(repo, answers=["", "", "y", ""])          # accept "commit locally without pushing"

    assert code == 1
    assert git(repo, "log", "-1", "--format=%s") != "remote commit"
    assert remote_log(repo)[0] == "remote commit"


# ---------------------------------------------------------------- standalone

def test_script_runs_on_its_own_without_the_rest_of_the_project(repo, tmp_path):
    """publish.bat only needs publish.py + utility/git_publisher.py + Python's standard library."""
    write(repo, "calculate/prezzi.py", "a = 2\n")
    project = Path(publish.__file__).resolve().parent
    shutil.copy(project / "publish.py", repo / "publish.py")
    (repo / "utility").mkdir()
    shutil.copy(project / "utility" / "__init__.py", repo / "utility" / "__init__.py")
    shutil.copy(project / "utility" / "git_publisher.py", repo / "utility" / "git_publisher.py")
    git(repo, "add", "-A")
    git(repo, "commit", "-qm", "tools")
    write(repo, "calculate/prezzi.py", "a = 3\n")

    env = {**os.environ, "PYTHONDONTWRITEBYTECODE": "1"}
    proc = subprocess.run([sys.executable, "publish.py", "--yes"], cwd=str(repo),
                          capture_output=True, text=True, encoding="utf-8", env=env)

    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "Done. Created commit" in proc.stdout
    assert remote_log(repo)[0] == "update(prezzi): update prezzi.py"
    assert not (repo / "logs").exists()                     # no app logging side effects


def test_publish_bat_is_ascii_with_windows_line_endings():
    raw = (Path(publish.__file__).resolve().parent / "publish.bat").read_bytes()
    raw.decode("ascii")
    assert raw.count(b"\r\n") == raw.count(b"\n")
    assert b"publish.py %*" in raw
