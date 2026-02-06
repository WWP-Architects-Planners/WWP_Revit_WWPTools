# -*- coding: utf-8 -*-

import os
import sys
import WWP_uiUtils as ui
from pyrevit import script

try:
    import subprocess
except Exception:
    subprocess = None

GITHUB_RELEASES_URL = "https://github.com/jason-svn/WWPTools/releases/latest"


def _find_repo_root(start_dir):
    cur = os.path.abspath(start_dir)
    while True:
        if os.path.isdir(os.path.join(cur, ".git")):
            return cur
        parent = os.path.dirname(cur)
        if parent == cur:
            return None
        cur = parent


def _run_git(args, repo_root):
    if subprocess is None:
        return 1, "subprocess unavailable"
    try:
        proc = subprocess.Popen(
            ["git"] + args,
            cwd=repo_root,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        out, err = proc.communicate()
        out = out.decode("utf-8", "ignore").strip()
        err = err.decode("utf-8", "ignore").strip()
        return proc.returncode, (out or err)
    except Exception as ex:
        return 1, str(ex)


repo_root = _find_repo_root(os.path.dirname(__file__))
if not repo_root:
    msg = "Could not find repo root (.git)."
    if ui.uiUtils_confirm(msg + "\n\nOpen download page?", title="Update WWPTools"):
        try:
            if subprocess is not None:
                subprocess.Popen(["explorer", GITHUB_RELEASES_URL])
        except Exception:
            pass
    raise SystemExit

code, _ = _run_git(["--version"], repo_root)
if code != 0:
    msg = "Git is not available on this machine."
    if ui.uiUtils_confirm(msg + "\n\nOpen download page?", title="Update WWPTools"):
        try:
            if subprocess is not None:
                subprocess.Popen(["explorer", GITHUB_RELEASES_URL])
        except Exception:
            pass
    raise SystemExit

code, err = _run_git(["remote", "get-url", "origin"], repo_root)
if code != 0:
    ui.uiUtils_alert("Git remote 'origin' is missing.\n\n{}".format(err), title="Update WWPTools")
    raise SystemExit

code, err = _run_git(["fetch", "origin", "main"], repo_root)
if code != 0:
    ui.uiUtils_alert("Failed to fetch updates.\n\n{}".format(err), title="Update WWPTools")
    raise SystemExit

code, counts = _run_git(["rev-list", "--left-right", "--count", "HEAD...origin/main"], repo_root)
if code != 0 or not counts:
    ui.uiUtils_alert("Failed to compare versions.\n\n{}".format(counts), title="Update WWPTools")
    raise SystemExit

behind = ahead = 0
try:
    left, right = counts.split()
    ahead = int(left)
    behind = int(right)
except Exception:
    ui.uiUtils_alert("Unexpected version compare output:\n{}".format(counts), title="Update WWPTools")
    raise SystemExit

if behind <= 0:
    if ahead > 0:
        ui.uiUtils_alert(
            "Your local copy is ahead of GitHub by {} commit(s).".format(ahead),
            title="Update WWPTools",
        )
    else:
        ui.uiUtils_alert("You are already running the latest version.", title="Update WWPTools")
    raise SystemExit

confirm = ui.uiUtils_confirm(
    "There are {} update(s) available.\n\nDo you want to update now?".format(behind),
    title="Update WWPTools",
)
if not confirm:
    raise SystemExit

code, msg = _run_git(["pull", "--ff-only", "origin", "main"], repo_root)
if code != 0:
    ui.uiUtils_alert("Update failed.\n\n{}".format(msg), title="Update WWPTools")
    raise SystemExit

ui.uiUtils_alert("Update complete. Please restart Revit.", title="Update WWPTools")
