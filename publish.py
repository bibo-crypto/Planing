"""Commit and push this project's local changes to GitHub.

Normally started through ``publish.bat`` (double-click). It looks at what
changed in this folder, shows the files, writes a commit message from the
changes, and after one confirmation commits and pushes.

    python publish.py              interactive (asks before committing)
    python publish.py --yes        no questions: commit every listed file and push
    python publish.py --dry-run    show what would be committed; change nothing
    python publish.py -m "text"    use "text" as the commit subject
    python publish.py --no-push    commit locally only

Safety: files ignored by .gitignore are never listed or committed; files that
look like secrets or are larger than 50 MB are left out unless you switch them
on; it never force-pushes, pulls or rewrites history. Logic lives in
utility/git_publisher.py.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Callable

from utility import git_publisher as gp

ROOT = Path(__file__).resolve().parent
_STATUS = {"A": "new     ", "M": "modified", "D": "deleted "}


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="publish", description="Commit and push this project's changes to GitHub.")
    parser.add_argument("-y", "--yes", action="store_true",
                        help="don't ask questions: commit every listed file and push")
    parser.add_argument("--dry-run", action="store_true",
                        help="show what would be committed and change nothing")
    parser.add_argument("--no-push", action="store_true", help="commit locally but don't push")
    parser.add_argument("--no-fetch", action="store_true",
                        help="don't contact GitHub to check for newer commits (offline)")
    parser.add_argument("-m", "--message", metavar="SUBJECT",
                        help="use this as the commit subject instead of the generated one")
    return parser.parse_args(argv)


def _lines_text(change: gp.Change) -> str:
    parts = []
    if change.added:
        parts.append(f"+{change.added}")
    if change.deleted:
        parts.append(f"-{change.deleted}")
    return " ".join(parts)


def _show_files(changes: list[gp.Change], selected: dict[str, bool], out: Callable[[str], None]) -> None:
    width = len(str(len(changes)))
    for number, change in enumerate(changes, start=1):
        mark = "[x]" if selected[change.path] else "[ ]"
        line = f"  {number:>{width}}. {mark} {_STATUS[change.status]}  {change.path}"
        if _lines_text(change):
            line += f"   ({_lines_text(change)})"
        out(line)
        if change.risk:
            out(f"  {' ' * width}      ^ WARNING: {change.risk}")


def _choose_files(changes: list[gp.Change], ask: Callable[[str], str],
                  out: Callable[[str], None]) -> dict[str, bool] | None:
    """Let the user toggle files; None means they quit."""
    selected = {c.path: not c.risk for c in changes}
    while True:
        out("")
        _show_files(changes, selected, out)
        out("")
        answer = ask("  Enter = continue | numbers to switch files on/off (e.g. 2,5) | q = quit: ").strip().lower()
        if answer in {"q", "quit"}:
            return None
        if not answer:
            return selected
        for token in answer.replace(",", " ").split():
            if token.isdigit() and 1 <= int(token) <= len(changes):
                path = changes[int(token) - 1].path
                selected[path] = not selected[path]
            else:
                out(f"  (ignored '{token}': not a file number)")


def _render(message: gp.CommitMessage) -> str:
    """Full commit text; a custom subject (empty type) is used as written."""
    subject = message.subject if message.type else message.summary
    return (subject + "\n\n" + message.body).strip() + "\n"


def _yes(answer: str, default: bool) -> bool:
    answer = answer.strip().lower()
    return default if not answer else answer in {"y", "yes"}


def main(argv: list[str] | None = None, *, ask: Callable[[str], str] = input,
         out: Callable[[str], None] = print, root: Path = ROOT) -> int:
    args = _parse_args(argv)
    interactive = not (args.yes or args.dry_run)

    def question(prompt: str) -> str:
        try:
            return ask(prompt)
        except EOFError:
            raise gp.GitError("No console input is available. Run `python publish.py --yes` to skip the questions.")

    try:
        out(f"  Folder: {root}")
        repo = gp.resolve_repo(root)
        info = gp.repo_info(repo, fetch=not args.no_fetch)
        target = f"{info.remote} ({info.remote_url})" if info.remote else "(no remote configured)"
        out(f"  Branch: {info.branch}  ->  {target}")
        if info.upstream:
            out(f"  Status: {info.ahead} commit(s) not pushed yet, {info.behind} commit(s) on GitHub not here")
        if info.fetch_error:
            out(f"  WARNING: couldn't check GitHub: {info.fetch_error}")
        changes = gp.list_changes(repo)
        out("")

        # -- nothing to commit: maybe only a push is pending ---------------------
        if not changes:
            if info.ahead and not args.no_push and not args.dry_run:
                out(f"  No file changes, but {info.ahead} commit(s) are waiting to be pushed.")
                if info.behind:
                    out(f"  GitHub has {info.behind} newer commit(s). Update first: git pull --rebase")
                    return 1
                if not info.remote:
                    out("  There is no remote to push to.")
                    return 1
                if interactive and not _yes(question(f"  Push them to {info.remote}/{info.branch}? [Y/n]: "), True):
                    out("  Cancelled. Nothing was pushed.")
                    return 0
                result = gp.publish(repo, [], "", do_push=True, progress=lambda t: out(f"  {t}"))
                out(f"  Done. Pushed {result.branch} to {result.remote_url or 'the remote'}.")
                return 0
            out("  Nothing to publish: no changes since the last commit.")
            return 0

        out(f"  {len(changes)} changed file(s)  (files ignored by .gitignore are not shown)")
        if interactive:
            selected = _choose_files(changes, question, out)
            if selected is None:
                out("  Cancelled. Nothing was committed.")
                return 0
        else:
            selected = {c.path: not c.risk for c in changes}
            out("")
            _show_files(changes, selected, out)
        chosen = [c for c in changes if selected[c.path]]
        if not chosen:
            out("")
            out("  No files selected. Nothing was committed.")
            return 0
        risky = [c for c in chosen if c.risk]
        if risky and interactive:
            out("")
            for change in risky:
                out(f"  WARNING: {change.path} - {change.risk}")
            if not _yes(question("  You switched on risky file(s). Really include them? [y/N]: "), False):
                out("  Cancelled. Nothing was committed.")
                return 0

        # -- commit message ------------------------------------------------------
        version = None
        if any(c.path == "version.txt" for c in chosen):
            try:
                version = (repo / "version.txt").read_text(encoding="utf-8").strip() or None
            except OSError:
                version = None
        message = gp.build_commit_message(chosen, version)
        if args.message:
            message.type, message.scope, message.summary = "", "", args.message.strip()
        out("")
        out("  Commit message:")
        for line in _render(message).splitlines():
            out(f"    {line}")
        if interactive and not args.message:
            own = question("\n  Enter = use this message | type a different subject line | q = quit: ").strip()
            if own.lower() in {"q", "quit"}:
                out("  Cancelled. Nothing was committed.")
                return 0
            if own:
                message.type, message.scope, message.summary = "", "", own
        full_message = _render(message)

        # -- can we push? ----------------------------------------------------------
        do_push = not args.no_push
        problem = ""
        if do_push and not info.remote:
            problem = "There is no remote to push to"
        elif do_push and info.behind:
            problem = f"GitHub has {info.behind} newer commit(s), so a push would be refused (update first: git pull --rebase)"
        if problem:
            out("")
            out(f"  WARNING: {problem}.")
            do_push = False
            if interactive and not _yes(question("  Commit locally without pushing? [y/N]: "), False):
                out("  Cancelled. Nothing was committed.")
                return 0

        if args.dry_run:
            out("")
            out(f"  Dry run: would commit {len(chosen)} file(s)" + (" and push." if do_push else ". (no push)"))
            return 0
        if interactive:
            where = f"commit {len(chosen)} file(s) and push to {info.remote}/{info.branch}" if do_push \
                else f"commit {len(chosen)} file(s) (no push)"
            if not _yes(question(f"\n  Ready to {where}. Continue? [Y/n]: "), True):
                out("  Cancelled. Nothing was committed.")
                return 0

        out("")
        result = gp.publish(repo, [c.path for c in chosen], full_message, do_push=do_push,
                            progress=lambda text: out(f"  {text}"))
        out("")
        out(f"  Done. Created commit {result.commit}"
            + (f" and pushed {result.branch} to {result.remote_url or 'the remote'}." if result.pushed else "."))
        return 1 if problem else 0

    except gp.PushFailedError as exc:
        out("")
        out("  " + str(exc).replace("\n", "\n  "))
        out("")
        out("  Your commit is safe on this computer. Fix the problem above and run publish again.")
        return 1
    except gp.GitError as exc:
        out("")
        out("  ERROR: " + str(exc).replace("\n", "\n  "))
        return 1
    except KeyboardInterrupt:
        out("")
        out("  Cancelled.")
        return 130


if __name__ == "__main__":
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")   # file names may contain Arabic
        except (AttributeError, ValueError):
            pass
    sys.exit(main())
