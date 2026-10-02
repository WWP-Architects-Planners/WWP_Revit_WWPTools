"""
Update service for WWPTools, modeled after LandscapeDataManager's UpdateService.
Checks GitHub for a newer release and hands the install to a separate PowerShell
window.  The extension's DLLs are locked while Revit runs, so that window
downloads the update while the user keeps working, waits for every Revit session
to close, then replaces the extension folder (ZIP path) or resets the git working
tree (git path).
"""

import json
import os
import re
import subprocess
import sys

_lib_path = os.path.dirname(os.path.abspath(__file__))
if _lib_path not in sys.path:
    sys.path.insert(0, _lib_path)

from WWP_compat import Request, urlopen, decode_to_text


AUTOMATIC_CHECK_INTERVAL_SECONDS = 86400  # 1 day

UPDATES_DIRECTORY = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
    "WWPTools", "Updates",
)
STATE_PATH = os.path.join(UPDATES_DIRECTORY, "update-state.json")

GITHUB_OWNER = "WWP-Architects-Planners"
GITHUB_REPO = "WWP_Revit_WWPTools"
GITHUB_API_URL = "https://api.github.com/repos/{}/{}/releases/latest".format(
    GITHUB_OWNER, GITHUB_REPO,
)
RELEASES_PAGE_URL = "https://github.com/{}/{}/releases/latest".format(
    GITHUB_OWNER, GITHUB_REPO,
)
REPO_CLONE_URL = "https://github.com/{}/{}.git".format(
    GITHUB_OWNER, GITHUB_REPO,
)

_scheduled_tag = None


# ---- Release model --------------------------------------------------------

class WWPRelease(object):
    """One published GitHub release of WWPTools."""

    def __init__(self, version, tag, name, notes, page_url, archive_url):
        self.version = version      # (major, minor, patch)
        self.tag = tag              # e.g. "V2.6.1"
        self.name = name            # release title
        self.notes = notes          # markdown body
        self.page_url = page_url    # html_url
        self.archive_url = archive_url


# ---- Version helpers ------------------------------------------------------

def parse_semver(value):
    """Parse 'v1.2.3', '1.2.3', 'V2.6.1' into (major, minor, patch)."""
    if not value:
        return None
    match = re.search(r"(\d+)\.(\d+)\.(\d+)", str(value))
    if not match:
        return None
    return tuple(int(x) for x in match.groups())


def get_installed_version():
    """Read the installed version from WWPTools.version.json."""
    try:
        path = os.path.join(_lib_path, "WWPTools.version.json")
        with open(path, "r") as f:
            data = json.load(f)
        return parse_semver(data.get("version") if isinstance(data, dict) else None)
    except Exception:
        return None


def installed_version_text():
    ver = get_installed_version()
    if ver:
        return "v{}.{}.{}".format(*ver)
    return "unknown"


def is_newer(available, installed):
    if not available or not installed:
        return False
    return available > installed


# ---- GitHub feed ----------------------------------------------------------

def parse_release(json_text):
    """Parse a GitHub releases/latest JSON response.  Returns WWPRelease or None."""
    try:
        data = json.loads(json_text)
    except Exception:
        return None
    if not isinstance(data, dict):
        return None
    if data.get("draft") or data.get("prerelease"):
        return None

    tag = data.get("tag_name") or ""
    version = parse_semver(tag)
    if not version:
        return None

    name = data.get("name") or tag
    notes = data.get("body") or ""
    page_url = data.get("html_url") or RELEASES_PAGE_URL
    archive_url = "https://github.com/{}/{}/archive/refs/tags/{}.zip".format(
        GITHUB_OWNER, GITHUB_REPO, tag,
    )
    return WWPRelease(version, tag, name, notes, page_url, archive_url)


def get_latest_release():
    """Fetch the latest release from GitHub.  Returns WWPRelease or None."""
    try:
        req = Request(GITHUB_API_URL, headers={
            "User-Agent": "WWPTools/{}".format(installed_version_text()),
            "Accept": "application/vnd.github+json",
        })
        resp = urlopen(req, timeout=15)
        text = decode_to_text(resp.read(), "utf-8")
        return parse_release(text)
    except Exception:
        return None


# ---- State persistence ----------------------------------------------------

def _read_state():
    try:
        if os.path.exists(STATE_PATH):
            with open(STATE_PATH, "r") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _write_state(state):
    try:
        if not os.path.isdir(UPDATES_DIRECTORY):
            os.makedirs(UPDATES_DIRECTORY)
        with open(STATE_PATH, "w") as f:
            json.dump(state, f)
    except Exception:
        pass


# ---- Background check -----------------------------------------------------

def check_in_background():
    """Startup check: at most once a day, skips versions the user dismissed.
    Returns WWPRelease or None.  Never throws."""
    try:
        state = _read_state()
        last_checked = state.get("last_checked_utc")
        if last_checked:
            try:
                from datetime import datetime
                last_dt = datetime.strptime(last_checked, "%Y-%m-%dT%H:%M:%S")
                elapsed = (datetime.utcnow() - last_dt).total_seconds()
                if elapsed < AUTOMATIC_CHECK_INTERVAL_SECONDS:
                    return None
            except Exception:
                pass

        release = get_latest_release()

        from datetime import datetime
        state["last_checked_utc"] = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S")
        _write_state(state)

        if release is None:
            return None

        installed = get_installed_version()
        if not is_newer(release.version, installed):
            return None

        skipped = state.get("skipped_tag")
        if skipped and skipped.lower() == release.tag.lower():
            return None

        return release
    except Exception:
        return None


def skip_version(release):
    """Persist the tag so this version won't prompt again on startup."""
    state = _read_state()
    state["skipped_tag"] = release.tag
    _write_state(state)


def is_scheduled(release):
    """True if this release's installer was already started this session."""
    global _scheduled_tag
    return (
        _scheduled_tag is not None
        and release is not None
        and _scheduled_tag.lower() == release.tag.lower()
    )


# ---- Installer generation -------------------------------------------------

def _find_extension_root():
    path = os.path.abspath(_lib_path)
    while path and path != os.path.dirname(path):
        if path.lower().endswith(".extension"):
            return path
        path = os.path.dirname(path)
    return os.path.normpath(os.path.join(_lib_path, ".."))


def _is_git_repo(path):
    return os.path.isdir(os.path.join(path, ".git"))


def _git_cli_available():
    try:
        p = subprocess.Popen(
            ["git", "--version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=False,
        )
        p.communicate()
        return p.returncode == 0
    except Exception:
        return False


def _current_git_branch(repo_root):
    try:
        p = subprocess.Popen(
            ["git", "-C", repo_root, "rev-parse", "--abbrev-ref", "HEAD"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=False,
        )
        stdout, _ = p.communicate()
        if p.returncode == 0:
            branch = (stdout or b"").decode("utf-8", "ignore").strip()
            if branch and branch != "HEAD":
                return branch
    except Exception:
        pass
    return "main"


_GIT_INSTALLER_SCRIPT = r"""
param(
    [Parameter(Mandatory)] [string] $Tag,
    [Parameter(Mandatory)] [string] $ExtensionRoot,
    [string] $Branch = 'main'
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'
$Host.UI.RawUI.WindowTitle = "WWPTools $Tag update"

try {
    Write-Host "Updating WWPTools to $Tag." -ForegroundColor Cyan
    Write-Host 'You can keep working in Revit while it downloads. Leave this window open.'
    Write-Host ''

    Write-Host 'Fetching update from GitHub...'
    git -C "$ExtensionRoot" fetch origin $Branch 2>&1
    if ($LASTEXITCODE -ne 0) { throw 'git fetch failed.' }
    Write-Host 'Fetch complete.'

    if (Get-Process -Name Revit -ErrorAction SilentlyContinue) {
        Write-Host ''
        Write-Host 'Save your work and close Revit (all open sessions) to install.' -ForegroundColor Yellow
        while (Get-Process -Name Revit -ErrorAction SilentlyContinue) {
            Start-Sleep -Seconds 2
        }
    }

    Write-Host 'Installing...'
    git -C "$ExtensionRoot" reset --hard "origin/$Branch" 2>&1
    if ($LASTEXITCODE -ne 0) { throw 'git reset failed.' }
    git -C "$ExtensionRoot" clean -ffdx 2>&1

    Write-Host ''
    Write-Host "WWPTools $Tag is installed. Start Revit to use it." -ForegroundColor Green
}
catch {
    Write-Host ''
    Write-Host "The update failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'If it stopped part-way through, re-run Update WWPTools from Revit or reinstall:'
    Write-Host "  https://github.com/WWP-ARCHITECTS-PLANNERS/WWP_Revit_WWPTools/releases/latest"
}

Write-Host ''
Read-Host 'Press Enter to close this window'
"""

_ZIP_INSTALLER_SCRIPT = r"""
param(
    [Parameter(Mandatory)] [string] $ArchiveUrl,
    [Parameter(Mandatory)] [string] $Tag,
    [Parameter(Mandatory)] [string] $ExtensionRoot,
    [Parameter(Mandatory)] [string] $WorkDirectory
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'
$Host.UI.RawUI.WindowTitle = "WWPTools $Tag update"

try {
    Write-Host "Updating WWPTools to $Tag." -ForegroundColor Cyan
    Write-Host 'You can keep working in Revit while it downloads. Leave this window open.'
    Write-Host ''

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    New-Item -ItemType Directory -Path $WorkDirectory -Force | Out-Null
    $zipPath     = Join-Path $WorkDirectory 'wwptools.zip'
    $extractPath = Join-Path $WorkDirectory 'extract'

    Write-Host 'Downloading the update from GitHub...'
    try {
        Start-BitsTransfer -Source $ArchiveUrl -Destination $zipPath `
            -DisplayName "WWPTools $Tag" -Description 'WWPTools update'
    }
    catch {
        # BITS can be disabled by policy; fall back to a plain download.
        (New-Object Net.WebClient).DownloadFile($ArchiveUrl, $zipPath)
    }

    Write-Host 'Extracting...'
    if (Test-Path -LiteralPath $extractPath) {
        Remove-Item -LiteralPath $extractPath -Recurse -Force
    }
    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractPath -Force

    $src = Get-ChildItem -LiteralPath $extractPath -Directory | Select-Object -First 1
    if (-not $src) { throw 'Downloaded ZIP did not contain a folder.' }

    if (Get-Process -Name Revit -ErrorAction SilentlyContinue) {
        Write-Host ''
        Write-Host 'Download complete. Save your work and close Revit (all open sessions) to install.' -ForegroundColor Yellow
        while (Get-Process -Name Revit -ErrorAction SilentlyContinue) {
            Start-Sleep -Seconds 2
        }
    }

    Write-Host 'Installing...'
    $extParent = Split-Path $ExtensionRoot -Parent
    if (-not (Test-Path $extParent)) {
        New-Item -ItemType Directory -Path $extParent -Force | Out-Null
    }
    if (Test-Path -LiteralPath $ExtensionRoot) {
        Remove-Item -LiteralPath $ExtensionRoot -Recurse -Force
    }
    Move-Item -LiteralPath $src.FullName -Destination $ExtensionRoot

    Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $extractPath) {
        Remove-Item -LiteralPath $extractPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    Write-Host ''
    Write-Host "WWPTools $Tag is installed. Start Revit to use it." -ForegroundColor Green
}
catch {
    Write-Host ''
    Write-Host "The update failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'If it stopped part-way through, reinstall WWPTools manually:'
    Write-Host "  https://github.com/WWP-ARCHITECTS-PLANNERS/WWP_Revit_WWPTools/releases/latest"
}

Write-Host ''
Read-Host 'Press Enter to close this window'
"""


def schedule_install(release):
    """Generate and launch the PowerShell installer for this release.
    Raises on failure."""
    global _scheduled_tag

    extension_root = os.path.normpath(_find_extension_root())
    use_git = _is_git_repo(extension_root) and _git_cli_available()

    if not os.path.isdir(UPDATES_DIRECTORY):
        os.makedirs(UPDATES_DIRECTORY)

    script_path = os.path.join(UPDATES_DIRECTORY, "Install-WWPToolsUpdate.ps1")

    if use_git:
        script_content = _GIT_INSTALLER_SCRIPT
    else:
        script_content = _ZIP_INSTALLER_SCRIPT

    # Write with BOM so Windows PowerShell 5.1 reads it as UTF-8.
    import codecs
    with codecs.open(script_path, "w", "utf-8-sig") as f:
        f.write(script_content)

    args = [
        "powershell.exe",
        "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", script_path,
    ]

    if use_git:
        branch = _current_git_branch(extension_root)
        args += [
            "-Tag", release.tag,
            "-ExtensionRoot", extension_root,
            "-Branch", branch,
        ]
    else:
        work_dir = os.path.join(UPDATES_DIRECTORY, release.tag)
        args += [
            "-ArchiveUrl", release.archive_url,
            "-Tag", release.tag,
            "-ExtensionRoot", extension_root,
            "-WorkDirectory", work_dir,
        ]

    try:
        subprocess.Popen(args, creationflags=0x00000010)  # CREATE_NEW_CONSOLE
        _scheduled_tag = release.tag
    except Exception as e:
        raise Exception("Could not start the update: {}".format(e))


# ---- TaskDialog prompt (mirrors LDM's UpdatePrompt) -----------------------

def _load_revit_ui():
    try:
        import clr
        clr.AddReference("RevitAPIUI")
        from Autodesk.Revit import UI
        return UI
    except Exception:
        return None


def show_update_prompt(release, offer_skip=False):
    """Show a Revit TaskDialog with Install / Remind / Skip options.
    Returns 'install', 'skip', or 'dismiss'."""
    if is_scheduled(release):
        UI = _load_revit_ui()
        if UI:
            UI.TaskDialog.Show(
                "WWPTools updates",
                "{} is already downloading in its own window. "
                "It installs when you close Revit.".format(release.tag),
            )
        return "dismiss"

    UI = _load_revit_ui()
    if UI is None:
        return "dismiss"

    notes = release.notes
    if len(notes) > 1500:
        notes = notes[:1500] + "..."

    dialog = UI.TaskDialog("WWPTools update available")
    dialog.MainInstruction = "{} is available".format(release.name)
    dialog.MainContent = (
        "You have {}. The update downloads in a separate window "
        "while you keep working, then installs once you close Revit."
    ).format(installed_version_text())
    if notes:
        dialog.ExpandedContent = notes
    dialog.FooterText = release.page_url
    dialog.CommonButtons = UI.TaskDialogCommonButtons.Close
    dialog.DefaultButton = UI.TaskDialogResult.Close

    dialog.AddCommandLink(
        UI.TaskDialogCommandLinkId.CommandLink1,
        "Download and install when Revit closes",
    )
    dialog.AddCommandLink(
        UI.TaskDialogCommandLinkId.CommandLink2,
        "Remind me later",
    )
    if offer_skip:
        dialog.AddCommandLink(
            UI.TaskDialogCommandLinkId.CommandLink3,
            "Skip {}".format(release.tag),
            "Don't remind me about this version again.",
        )

    result = dialog.Show()

    if result == UI.TaskDialogResult.CommandLink1:
        try:
            schedule_install(release)
        except Exception as e:
            UI.TaskDialog.Show(
                "WWPTools updates",
                "Could not start the update: {}".format(e),
            )
        return "install"

    if result == UI.TaskDialogResult.CommandLink3:
        skip_version(release)
        return "skip"

    return "dismiss"


def show_up_to_date():
    """Show a dialog confirming the user has the latest version."""
    UI = _load_revit_ui()
    if UI:
        UI.TaskDialog.Show(
            "WWPTools updates",
            "You have the latest version of WWPTools ({}).".format(
                installed_version_text(),
            ),
        )


def show_check_failed(error_message=None):
    """Show a dialog when the update check fails (manual check only)."""
    UI = _load_revit_ui()
    if UI:
        msg = "Could not reach GitHub to check for updates"
        if error_message:
            msg += " ({})".format(error_message)
        msg += ".\n\nYou have {}. Releases: {}".format(
            installed_version_text(), RELEASES_PAGE_URL,
        )
        UI.TaskDialog.Show("WWPTools updates", msg)
