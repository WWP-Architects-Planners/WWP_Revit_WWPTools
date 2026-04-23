import os
import subprocess
import sys
import traceback

from pyrevit import script
from pyrevit.coreutils import git as pygit


script_dir = os.path.dirname(__file__)
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

import WWP_uiUtils as ui


TITLE = "Update WWPTools"
RELEASES_URL = "https://github.com/WWP-Architects-Planners/WWP_Revit_WWPTools/releases/latest"
TARGET_BRANCH = "main"
TARGET_REMOTE_BRANCH = "origin/main"


def _extension_root():
    return os.path.normpath(os.path.join(script_dir, "..", "..", ".."))


def _discover_repo(extension_root):
    try:
        repo_path = pygit.libgit.Repository.Discover(extension_root)
    except Exception:
        return None
    if not repo_path:
        return None
    try:
        return pygit.get_repo(repo_path)
    except Exception:
        return None


def _git_cli_available():
    try:
        completed = subprocess.Popen(
            ["git", "--version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=False,
        )
        completed.communicate()
        return completed.returncode == 0
    except Exception:
        return False


def _run_git(repo_root, args):
    completed = subprocess.Popen(
        ["git", "-C", repo_root] + list(args),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        shell=False,
    )
    stdout, stderr = completed.communicate()
    if completed.returncode != 0:
        raise Exception((stderr or stdout or b"Git command failed.").decode("utf-8", "ignore").strip())


def _git_output(repo_root, args):
    completed = subprocess.Popen(
        ["git", "-C", repo_root] + list(args),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        shell=False,
    )
    stdout, stderr = completed.communicate()
    if completed.returncode != 0:
        raise Exception((stderr or stdout or b"Git command failed.").decode("utf-8", "ignore").strip())
    return (stdout or b"").decode("utf-8", "ignore")


def _get_dirty_paths(repo_root):
    if not _git_cli_available():
        return []
    output = _git_output(repo_root, ["status", "--porcelain"])
    dirty = []
    for line in output.splitlines():
        if not line.strip():
            continue
        path = line[3:] if len(line) > 3 else line
        dirty.append(path)
    return dirty


def _show_dirty_repo_message(repo_info, dirty_paths):
    preview = dirty_paths[:12]
    lines = [
        "WWPTools can not auto-update this install because the local repo has uncommitted changes.",
        "",
        "Current branch: {}".format(repo_info.branch),
        "Target branch: {}".format(TARGET_REMOTE_BRANCH),
        "",
        "Dirty files:",
    ]
    lines.extend(preview)
    if len(dirty_paths) > len(preview):
        lines.append("... and {} more".format(len(dirty_paths) - len(preview)))
    lines.extend([
        "",
        "Commit or stash these changes first, then run Update WWPTools again.",
        "",
        "Open the latest GitHub release instead?",
    ])

    if ui.uiUtils_confirm("\n".join(lines), TITLE):
        _open_latest_release()


def _ensure_target_branch(repo_info, repo_root):
    if repo_info.branch == TARGET_BRANCH:
        return repo_info

    dirty_paths = _get_dirty_paths(repo_root)
    if dirty_paths:
        _show_dirty_repo_message(repo_info, dirty_paths)
        return None

    repo = repo_info.repo
    try:
        target_branch = repo.Branches[TARGET_BRANCH]
    except Exception:
        target_branch = None

    if target_branch is not None:
        try:
            pygit.libgit.Commands.Checkout(repo, target_branch)
            return pygit.get_repo(repo_root)
        except Exception:
            pass

    if not _git_cli_available():
        raise Exception(
            "Installed repo is on branch '{}', but this updater is configured to use '{}'.\n\n"
            "Git CLI is not available to switch branches automatically.".format(
                repo_info.branch,
                TARGET_REMOTE_BRANCH,
            )
        )

    _run_git(repo_root, ["fetch", "origin", TARGET_BRANCH])
    _run_git(repo_root, ["checkout", "-B", TARGET_BRANCH, TARGET_REMOTE_BRANCH])
    return pygit.get_repo(repo_root)


class _DivergenceResult(object):
    def __init__(self, behind, ahead):
        self.BehindBy = behind
        self.AheadBy = ahead


def _history_divergence(repo_info, repo_root):
    try:
        pygit.git_fetch(repo_info)
        return pygit.compare_branch_heads(repo_info)
    except Exception:
        pass
    if not _git_cli_available():
        return None
    try:
        _run_git(repo_root, ["fetch", "origin", TARGET_BRANCH])
        behind = int(_git_output(repo_root, ["rev-list", "--count", "HEAD..origin/{}".format(TARGET_BRANCH)]).strip())
        ahead = int(_git_output(repo_root, ["rev-list", "--count", "origin/{}..HEAD".format(TARGET_BRANCH)]).strip())
        return _DivergenceResult(behind, ahead)
    except Exception:
        return None


def _open_latest_release():
    script.open_url(RELEASES_URL)


def _show_not_repo_message():
    open_release = ui.uiUtils_confirm(
        "This installation is not a Git clone, so WWPTools can not update it with built-in Git.\n\n"
        "Open the latest GitHub release instead?",
        TITLE,
    )
    if open_release:
        _open_latest_release()


def _update_repo(repo_info, repo_root):
    repo_info = _ensure_target_branch(repo_info, repo_root)
    if repo_info is None:
        return
    divergence = _history_divergence(repo_info, repo_root)
    behind = int(divergence.BehindBy) if divergence and divergence.BehindBy is not None else 0
    ahead = int(divergence.AheadBy) if divergence and divergence.AheadBy is not None else 0

    if behind <= 0:
        msg = "WWPTools is already up to date on branch '{}'.\nCurrent commit: {}".format(
            repo_info.branch,
            repo_info.last_commit_hash[:7],
        )
        if ahead > 0:
            msg += "\n\nLocal repo is ahead of remote by {} commit(s).".format(ahead)
        ui.uiUtils_alert(msg, TITLE)
        return

    confirm = ui.uiUtils_confirm(
        "Updates are available for WWPTools.\n\n"
        "Branch: {}\n"
        "Remote: {}\n"
        "Current commit: {}\n"
        "Incoming commits: {}\n\n"
        "Update now using Git support from this installed clone?".format(
            repo_info.branch,
            TARGET_REMOTE_BRANCH,
            repo_info.last_commit_hash[:7],
            behind,
        ),
        TITLE,
    )
    if not confirm:
        return

    before_hash = repo_info.last_commit_hash[:7]
    try:
        updated_repo = pygit.git_pull(repo_info)
    except Exception:
        if not _git_cli_available():
            raise
        _run_git(repo_root, ["pull", "--ff-only", "origin", TARGET_BRANCH])
        updated_repo = pygit.get_repo(repo_root)
    after_hash = updated_repo.last_commit_hash[:7]

    ui.uiUtils_alert(
        "WWPTools updated successfully.\n\n"
        "Branch: {}\n"
        "Previous commit: {}\n"
        "Current commit: {}\n\n"
        "Restart Revit or reload pyRevit to ensure all changes are picked up.".format(
            updated_repo.branch,
            before_hash,
            after_hash,
        ),
        TITLE,
    )


def main():
    repo_root = _extension_root()
    repo_info = _discover_repo(repo_root)
    if not repo_info:
        _show_not_repo_message()
        return
    _update_repo(repo_info, repo_root)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        ui.uiUtils_alert(traceback.format_exc(), TITLE)
