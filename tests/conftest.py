import pytest

from tests.git_helpers import git, write


@pytest.fixture
def repo(tmp_path):
    """A repo with one commit, a .gitignore and a bare 'origin' it is already tracking."""
    remote = tmp_path / "remote.git"
    work = tmp_path / "work"
    git(tmp_path, "init", "--bare", "-b", "main", str(remote))
    git(tmp_path, "init", "-b", "main", str(work))
    git(work, "config", "user.name", "Test User")
    git(work, "config", "user.email", "test@example.com")
    write(work, ".gitignore", "settings/\n*.sqlite3\n*.xlsx\n")
    write(work, "calculate/prezzi.py", "a = 1\n")
    write(work, "gui/gui.py", "b = 1\n")
    git(work, "add", "-A")
    git(work, "commit", "-m", "initial")
    git(work, "remote", "add", "origin", str(remote))
    git(work, "push", "-u", "origin", "main")
    return work
