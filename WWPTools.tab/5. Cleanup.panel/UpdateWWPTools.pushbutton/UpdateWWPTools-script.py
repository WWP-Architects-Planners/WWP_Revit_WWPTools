import os
import subprocess
import sys
import tempfile
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

# Windows process-creation flags (safe fallback for IronPython)
_DETACHED_PROCESS       = getattr(subprocess, "DETACHED_PROCESS",       0x00000008)
_CREATE_NEW_PROCESS_GROUP = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)


def _reload_pyrevit():
    try:
        from pyrevit.loader import sessionmgr
        sessionmgr.reload_pyrevit()
        return True
    except Exception:
        pass
    try:
        from pyrevit.loader import sessionmgr as sm
        sm.reload_pyrevit()
        return True
    except Exception:
        return False


def _extension_root():
    return os.path.normpath(os.path.join(script_dir, "..", "..", ".."))


def _latest_tag(repo_root):
    if not _git_cli_available():
        return None
    try:
        tag = _git_output(repo_root, ["describe", "--tags", "--abbrev=0"]).strip()
        return tag if tag else None
    except Exception:
        return None


def _remote_tag(repo_root):
    if not _git_cli_available():
        return None
    try:
        tag = _git_output(
            repo_root,
            ["describe", "--tags", "--abbrev=0", "origin/{}".format(TARGET_BRANCH)]
        ).strip()
        return tag if tag else None
    except Exception:
        return None


def _incoming_log(repo_root, max_lines=10):
    if not _git_cli_available():
        return ""
    try:
        log = _git_output(
            repo_root,
            ["log", "--oneline", "HEAD..origin/{}".format(TARGET_BRANCH)]
        ).strip()
        lines = log.splitlines()
        if len(lines) > max_lines:
            lines = lines[:max_lines] + ["... and {} more".format(len(lines) - max_lines)]
        return "\n".join(lines)
    except Exception:
        return ""


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



def _ensure_target_branch(repo_info, repo_root):
    if repo_info.branch == TARGET_BRANCH:
        return repo_info

    # Don't block on dirty files — we always reset --hard so local changes are discarded anyway.
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
        ahead  = int(_git_output(repo_root, ["rev-list", "--count", "origin/{}..HEAD".format(TARGET_BRANCH)]).strip())
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


# ---------------------------------------------------------------------------
# Background updater — runs after Revit exits
# ---------------------------------------------------------------------------

def _schedule_update_on_revit_exit(repo_root):
    """
    Spawn a detached cmd process that waits until this Revit PID disappears,
    then runs git pull so locked DLLs are no longer an obstacle.
    Returns True if the watcher was started successfully.
    """
    pid = os.getpid()
    safe_root = os.path.normpath(repo_root)

    # Batch script: poll until the PID is gone, then pull
    batch = "\r\n".join([
        "@echo off",
        "title WWPTools — waiting for Revit to close...",
        ":wait",
        "tasklist /FI \"PID eq {pid}\" 2>NUL | find \"{pid}\" >NUL".format(pid=pid),
        "if not errorlevel 1 (timeout /t 3 /nobreak >NUL & goto wait)",
        "title WWPTools — applying update...",
        "echo.",
        "echo Revit closed.  Pulling latest WWPTools...",
        "echo.",
        "git -C \"{root}\" fetch origin {branch}".format(root=safe_root, branch=TARGET_BRANCH),
        "git -C \"{root}\" reset --hard origin/{branch}".format(root=safe_root, branch=TARGET_BRANCH),
        "if errorlevel 1 (",
        "  echo.",
        "  echo Update failed.  Open Revit and run Update WWPTools to try again.",
        "  pause",
        ") else (",
        "  echo.",
        "  echo WWPTools updated successfully.",
        "  timeout /t 5 /nobreak >NUL",
        ")",
    ]) + "\r\n"

    batch_path = os.path.join(
        tempfile.gettempdir(),
        "wwptools_update_{}.bat".format(pid),
    )
    try:
        with open(batch_path, "w") as f:
            f.write(batch)
        subprocess.Popen(
            ["cmd", "/c", batch_path],
            creationflags=_DETACHED_PROCESS | _CREATE_NEW_PROCESS_GROUP,
            close_fds=True,
        )
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Main update flow
# ---------------------------------------------------------------------------

def _update_repo(repo_info, repo_root):
    repo_info = _ensure_target_branch(repo_info, repo_root)
    if repo_info is None:
        return
    divergence = _history_divergence(repo_info, repo_root)
    behind = int(divergence.BehindBy) if divergence and divergence.BehindBy is not None else 0
    ahead  = int(divergence.AheadBy)  if divergence and divergence.AheadBy  is not None else 0

    current_tag = _latest_tag(repo_root)
    current_label = "{} ({})".format(current_tag, repo_info.last_commit_hash[:7]) if current_tag \
        else repo_info.last_commit_hash[:7]

    if behind <= 0:
        msg = "WWPTools is already up to date.\n\nVersion: {}\nBranch: {}".format(
            current_label, repo_info.branch,
        )
        if ahead > 0:
            msg += "\n\nLocal repo is {} commit(s) ahead of remote.".format(ahead)
        ui.uiUtils_alert(msg, TITLE)
        return

    remote_tag = _remote_tag(repo_root)
    remote_label = "{} ({})".format(remote_tag, "incoming") if remote_tag else "{} commit(s)".format(behind)
    changelog = _incoming_log(repo_root)

    confirm_msg = (
        "Updates are available for WWPTools.\n\n"
        "Current version:  {}\n"
        "New version:      {}\n"
        "Branch:           {}\n\n"
        "What's new:\n{}\n\n"
        "Update now?"
    ).format(
        current_label,
        remote_label,
        repo_info.branch,
        changelog if changelog else "  (commit log unavailable)",
    )
    if not ui.uiUtils_confirm(confirm_msg, TITLE):
        return

    before_hash = repo_info.last_commit_hash[:7]
    dll_locked = False
    try:
        updated_repo = pygit.git_pull(repo_info)
    except Exception:
        if not _git_cli_available():
            raise
        try:
            _run_git(repo_root, ["fetch", "origin", TARGET_BRANCH])
            _run_git(repo_root, ["reset", "--hard", "origin/{}".format(TARGET_BRANCH)])
        except Exception as pull_err:
            err_str = str(pull_err)
            if "unable to unlink" in err_str or ("unlink" in err_str and "dll" in err_str.lower()):
                dll_locked = True
            else:
                raise
        if not dll_locked:
            updated_repo = pygit.get_repo(repo_root)

    if dll_locked:
        if _schedule_update_on_revit_exit(repo_root):
            ui.uiUtils_alert(
                "WWPTools will update automatically when you close Revit.\n\n"
                "A small console window will appear after Revit exits to confirm.\n"
                "You do not need to do anything else.",
                TITLE,
            )
        else:
            ui.uiUtils_alert(
                "Could not schedule the background update.\n\n"
                "To update manually: close Revit completely, then run Update WWPTools again.",
                TITLE,
            )
        return

    after_hash = updated_repo.last_commit_hash[:7]
    new_tag = _latest_tag(repo_root)
    after_label = "{} ({})".format(new_tag, after_hash) if new_tag else after_hash

    reload_offered = ui.uiUtils_confirm(
        "WWPTools updated successfully.\n\n"
        "Previous version: {}\n"
        "New version:      {}\n\n"
        "Reload pyRevit now to apply the changes?\n"
        "(Choosing No means you'll need to restart Revit manually.)".format(
            current_label,
            after_label,
        ),
        TITLE,
    )
    if reload_offered:
        if not _reload_pyrevit():
            ui.uiUtils_alert(
                "Could not reload pyRevit automatically.\n\nPlease restart Revit to apply the update.",
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
