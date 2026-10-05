"""Tests for utility.git_publisher, run against real throw-away git repositories."""
import pytest

from tests.git_helpers import git, remote_log, write
from utility import git_publisher as gp

pytestmark = pytest.mark.skipif(gp.find_git() is None, reason="git is not installed")


# ---------------------------------------------------------------- status

def test_ignored_files_never_appear_and_are_never_committed(repo):
    write(repo, "settings/prefs.json", "{}")
    write(repo, "data/orders.sqlite3", "db")
    write(repo, "report.xlsx", "xlsx")
    write(repo, "calculate/new_module.py", "n = 1\n")

    changes = gp.list_changes(repo)

    assert [c.path for c in changes] == ["calculate/new_module.py"]
    gp.publish(repo, ["calculate/new_module.py"], "feat: add module", do_push=False)
    tracked = git(repo, "ls-files").splitlines()
    assert "settings/prefs.json" not in tracked and "report.xlsx" not in tracked
    assert "data/orders.sqlite3" not in tracked


def test_list_changes_reports_status_and_line_counts(repo):
    write(repo, "calculate/prezzi.py", "a = 1\nb = 2\nc = 3\n")      # modified: +2
    (repo / "gui" / "gui.py").unlink()                                 # deleted: -1
    write(repo, "utility/new.py", "1\n2\n3\n4\n")                      # new: 4 lines

    by_path = {c.path: c for c in gp.list_changes(repo)}

    assert by_path["calculate/prezzi.py"].status == "M" and by_path["calculate/prezzi.py"].added == 2
    assert by_path["gui/gui.py"].status == "D" and by_path["gui/gui.py"].deleted == 1
    assert by_path["utility/new.py"].status == "A" and by_path["utility/new.py"].added == 4


def test_risky_files_are_flagged(repo, monkeypatch):
    write(repo, ".env", "TOKEN=abc")
    write(repo, "keys/server.pem", "pem")
    write(repo, "calculate/ok.py", "x = 1\n")
    monkeypatch.setattr(gp, "LARGE_FILE_BYTES", 10)
    write(repo, "data/big.bin", "0123456789ABCDEF")

    risks = {c.path: c.risk for c in gp.list_changes(repo)}

    assert "secret" in risks[".env"] and "secret" in risks["keys/server.pem"]
    assert "larger than" in risks["data/big.bin"]
    assert risks["calculate/ok.py"] == ""


def test_unresolved_merge_conflict_is_refused(repo):
    git(repo, "checkout", "-q", "-b", "other")
    write(repo, "gui/gui.py", "b = 'other'\n"); git(repo, "commit", "-qam", "other")
    git(repo, "checkout", "-q", "main")
    write(repo, "gui/gui.py", "b = 'main'\n"); git(repo, "commit", "-qam", "main")
    git(repo, "merge", "other", check=False)

    with pytest.raises(gp.GitError, match="conflict"):
        gp.list_changes(repo)


# ---------------------------------------------------------------- repo info

def test_repo_info_reports_branch_remote_and_ahead_behind(repo):
    write(repo, "a.py"); git(repo, "add", "-A"); git(repo, "commit", "-qm", "local only")

    info = gp.repo_info(repo)

    assert info.branch == "main" and info.remote == "origin" and info.upstream == "origin/main"
    assert (info.ahead, info.behind) == (1, 0)
    assert info.fetch_error == ""


def test_credentials_in_the_remote_url_are_masked():
    assert gp.mask_url("https://user:ghp_secret@github.com/o/r.git") == "https://github.com/o/r.git"
    assert gp.mask_url("https://github.com/o/r.git") == "https://github.com/o/r.git"
    assert gp.mask_url("git@github.com:o/r.git") == "git@github.com:o/r.git"


def test_detached_head_is_refused(repo):
    git(repo, "checkout", "-q", "--detach")
    with pytest.raises(gp.GitError, match="detached"):
        gp.repo_info(repo)


def test_resolve_repo_finds_the_root_and_rejects_other_folders(repo, tmp_path):
    assert gp.resolve_repo(repo / "gui") == repo.resolve() or gp.resolve_repo(repo / "gui").samefile(repo)
    plain = tmp_path / "plain"
    plain.mkdir()
    with pytest.raises(gp.GitError, match="not inside a Git repository"):
        gp.resolve_repo(plain)


# ---------------------------------------------------------------- publish

def test_publish_commits_only_the_selected_files_and_pushes(repo):
    write(repo, "calculate/prezzi.py", "a = 2\n")
    write(repo, "gui/gui.py", "b = 2\n")           # changed but NOT selected
    write(repo, "utility/other.py", "o = 1\n")     # new, NOT selected
    git(repo, "add", "utility/other.py")           # already staged by the user: must stay out of the commit

    result = gp.publish(repo, ["calculate/prezzi.py"], "update(prezzi): tweak", do_push=True)

    assert result.pushed and result.commit and result.branch == "main"
    assert git(repo, "show", "--name-only", "--format=", "HEAD").splitlines() == ["calculate/prezzi.py"]
    assert remote_log(repo)[0] == "update(prezzi): tweak"
    remaining = {c.path for c in gp.list_changes(repo)}
    assert remaining == {"gui/gui.py", "utility/other.py"}
    assert "utility/other.py" in git(repo, "diff", "--cached", "--name-only")   # still staged, untouched


def test_publish_handles_new_modified_and_deleted_files_together(repo):
    write(repo, "calculate/prezzi.py", "a = 9\n")
    (repo / "gui" / "gui.py").unlink()
    write(repo, "utility/git_x.py", "z = 1\n")

    gp.publish(repo, ["calculate/prezzi.py", "gui/gui.py", "utility/git_x.py"], "chore: mixed")

    assert gp.list_changes(repo) == []
    files = git(repo, "show", "--name-status", "--format=", "HEAD").splitlines()
    assert sorted(f.split("\t")[0] for f in files) == ["A", "D", "M"]


def test_publish_with_unicode_paths_and_message(repo):
    write(repo, "docs/ملاحظات.md", "مرحبا\n")

    gp.publish(repo, ["docs/ملاحظات.md"], "docs: ملاحظات جديدة\n\n- A docs/ملاحظات.md")

    assert remote_log(repo)[0] == "docs: ملاحظات جديدة"


def test_many_paths_do_not_hit_command_line_limits(repo):
    names = [f"gen/file_{i:04d}_{'x' * 80}.py" for i in range(300)]
    for name in names:
        write(repo, name)

    gp.publish(repo, names, "chore: many files")

    assert len(git(repo, "show", "--name-only", "--format=", "HEAD").splitlines()) == 300


def test_first_push_sets_the_upstream(repo):
    git(repo, "checkout", "-q", "-b", "feature")
    write(repo, "f.py")

    result = gp.publish(repo, ["f.py"], "feat: f")

    assert result.pushed
    assert git(repo, "rev-parse", "--abbrev-ref", "@{u}") == "origin/feature"


def test_push_only_sends_commits_that_already_exist_locally(repo):
    write(repo, "x.py"); git(repo, "add", "-A"); git(repo, "commit", "-qm", "already committed")

    result = gp.publish(repo, [], "", do_push=True)

    assert result.pushed and result.commit == ""
    assert remote_log(repo)[0] == "already committed"


def test_rejected_push_keeps_the_local_commit_and_explains(repo, tmp_path):
    other = tmp_path / "other"
    git(tmp_path, "clone", "-q", git(repo, "remote", "get-url", "origin"), str(other))
    git(other, "config", "user.name", "B"); git(other, "config", "user.email", "b@example.com")
    write(other, "remote_only.py"); git(other, "add", "-A"); git(other, "commit", "-qm", "remote commit")
    git(other, "push", "-q")

    write(repo, "calculate/prezzi.py", "a = 3\n")
    with pytest.raises(gp.PushFailedError) as caught:
        gp.publish(repo, ["calculate/prezzi.py"], "update: mine")

    assert "pull --rebase" in str(caught.value) and caught.value.commit
    assert git(repo, "log", "-1", "--format=%s") == "update: mine"       # commit kept locally
    assert "update: mine" not in remote_log(repo)                        # and nothing was forced
    assert gp.repo_info(repo).behind == 1


def test_no_remote_gives_a_clear_error(tmp_path):
    work = tmp_path / "solo"
    git(tmp_path, "init", "-b", "main", str(work))
    git(work, "config", "user.name", "T"); git(work, "config", "user.email", "t@example.com")
    write(work, "a.py")

    with pytest.raises(gp.PushFailedError, match="no remote"):
        gp.publish(work, ["a.py"], "feat: first")
    assert git(work, "log", "-1", "--format=%s") == "feat: first"       # initial commit still made


def test_empty_message_and_unchanged_selection_are_refused(repo):
    write(repo, "a.py")
    with pytest.raises(gp.GitError, match="message is empty"):
        gp.publish(repo, ["a.py"], "   ")
    with pytest.raises(gp.GitError, match="nothing to commit"):
        gp.publish(repo, ["gui/gui.py"], "chore: nothing")              # unchanged file


def test_progress_messages_are_reported(repo):
    write(repo, "a.py")
    steps = []
    gp.publish(repo, ["a.py"], "feat: a", progress=steps.append)
    assert any("Staging" in s for s in steps) and any("Pushing" in s for s in steps)


def test_error_texts_are_translated():
    assert "pull --rebase" in gp.explain_git_error("! [rejected] main -> main (fetch first)")
    assert "credentials" in gp.explain_git_error("fatal: Authentication failed for 'https://github.com/x'")
    assert "user.email" in gp.explain_git_error("fatal: unable to auto-detect email address")
    assert "internet" in gp.explain_git_error("fatal: unable to access: Could not resolve host: github.com")


# ---------------------------------------------------------------- message

def _c(path, status="M", added=3, deleted=1):
    return gp.Change(path=path, status=status, added=added, deleted=deleted)


def test_area_names():
    assert gp.area_of("calculate/prezzi.py") == "prezzi"
    assert gp.area_of("gui/tabs/prezzi_tab.py") == "prezzi"
    assert gp.area_of("gui/tabs/master_data_tab.py") == "master-data"
    assert gp.area_of("exporters/biglietti_exporter.py") == "biglietti"
    assert gp.area_of("gui/gui.py") == "gui"
    assert gp.area_of("tests/test_x.py") == "tests"
    assert gp.area_of(".github/workflows/tests.yml") == "ci"
    assert gp.area_of("README.md") == "docs"
    assert gp.area_of("requirements.txt") == "build"


def test_message_for_a_single_modified_file():
    message = gp.build_commit_message([_c("calculate/prezzi.py", added=12, deleted=4)])
    assert message.subject == "update(prezzi): update prezzi.py"
    assert "- M calculate/prezzi.py (+12/-4)" in message.body


def test_message_for_new_source_file_is_a_feature():
    message = gp.build_commit_message([
        _c("utility/git_publisher.py", "A", 300, 0),
        _c("gui/gui.py", "M", 8, 1),
        _c("tests/test_git_publisher.py", "A", 100, 0),
    ])
    assert message.type == "feat"
    assert message.scope.split(", ")[0] == "git-publisher" and "tests" not in message.scope
    assert message.summary.startswith("add git_publisher.py")
    assert "- A utility/git_publisher.py (+300)" in message.body


def test_message_for_tests_docs_ci_and_build_only():
    assert gp.build_commit_message([_c("tests/test_a.py"), _c("tests/test_b.py")]).type == "test"
    assert gp.build_commit_message([_c("README.md")]).type == "docs"
    assert gp.build_commit_message([_c(".github/workflows/tests.yml", "A")]).type == "ci"
    assert gp.build_commit_message([_c("requirements.txt"), _c("main.spec")]).type == "build"


def test_message_for_release_uses_the_version():
    message = gp.build_commit_message([_c("version.txt", added=1, deleted=1), _c("utility/updater.py", added=1, deleted=1)],
                                      version="1.1.4")
    assert message.subject == "chore(release): bump version to 1.1.4"
    assert "utility/updater.py" in message.body


def test_message_for_many_files_summarises_counts_and_lines():
    message = gp.build_commit_message([
        _c("calculate/prezzi.py", "M", 10, 2), _c("gui/gui.py", "M", 5, 5), _c("gui/old.py", "D", 0, 20),
    ])
    assert message.type == "update"
    assert message.summary == "2 modified, 1 deleted (+15/-27)"


def test_long_file_lists_are_truncated_in_the_body():
    message = gp.build_commit_message([_c(f"gen/f{i}.py", "A", 1, 0) for i in range(55)])
    assert message.body.splitlines()[-1] == "- ... and 15 more files"


def test_full_message_has_subject_blank_line_and_body():
    message = gp.CommitMessage("fix", "prezzi", "handle ties", "- M calculate/prezzi.py")
    assert message.full() == "fix(prezzi): handle ties\n\n- M calculate/prezzi.py\n"
    assert gp.CommitMessage("fix", "", "x", "").full() == "fix: x\n"
