# WWPTools (pyRevit Extension)

WWPTools is a pyRevit toolbar extension distributed via GitHub Releases for easy installs and updates.

## pyRevit Extension Manager
pyRevit installs a Git extension by cloning the repository root directly into `%APPDATA%\pyRevit\Extensions\WWPTools.extension`.

That means `WWPTools.extension` can not live as a nested folder inside the same branch that users install from. If the repo root contains `backend`, `installer`, `WWPTools.WpfUI`, and other development folders, pyRevit clones all of them into the installed extension folder and the toolbar only works when the actual tab/panel bundles are also at that root.

This repo now supports a two-branch model:

1. `main` stays the development branch and keeps source projects such as `WWPTools.WpfUI`, `backend`, `installer`, and solution files.
2. `pyrevit` is auto-published from the `WWPTools.extension` subtree, so its branch root is the extension payload itself.

### Required GitHub setup
1. Push `main` as usual.
2. Let the `Publish pyRevit Branch` workflow create/update the `pyrevit` branch.
3. Set the repository default branch to `pyrevit` if you want pyRevit Extension Manager installs to pull the clean extension-only layout.

If you do not want to change the default branch, use a separate distribution repo that mirrors the `pyrevit` branch.

## For admins (publish updates)
1) Create a GitHub repo named `WWPTools` (public).
2) Publish updates by pushing to `main` (the installer pulls the repo zip from GitHub).
3) If the repo owner or name changes, update it in `installer/WWPTools.wxs` and rebuild the MSI.
4) Build the MSI by running `installer/Build_WWPTools_MSI.ps1`.
5) Attach the MSI to a GitHub release (example: `V1.0.0`).

## For users (install or update)
1) Download the MSI from the latest GitHub release (example: `V1.0.0`).
2) Run `WWPTools.msi`.
3) The installer shows the install path and lets you change it if needed.
4) The installer installs/updates the extension in:
   `%APPDATA%\pyRevit\Extensions\WWPTools.extension`
   and installs Dynamo packages to `C:\dynpackages`.

## Dynamo packages
The installer adds `C:\dynpackages` to Dynamo's Custom Package Folders (for any
installed Dynamo Revit versions), so users keep their personal packages too. If
Dynamo has never been launched, the installer creates the settings file and will
prompt users to open Dynamo once to finalize setup.

## Version display
WWPTools reads its installed version from `WWPTools.extension/lib/WWPTools.version.json`.
Tool dialogs append that version to their window titles so users can confirm what
build they are running without leaving the tool.

## Usage telemetry and weekly reports
WWPTools now includes an optional telemetry client plus a small self-hostable
backend in `backend/wwptools_usage_server.py`.

1. Start the backend:
   `python backend/wwptools_usage_server.py --host 0.0.0.0 --port 8787 --db data/wwptools-usage.db --api-key your-secret`
2. Copy `backend/telemetry.client.sample.json` to `%APPDATA%\pyRevit\WWPTools\telemetry\telemetry.config.json` on client machines and set the real `endpoint_url` and `api_key`.
3. The add-in records anonymized `app-init` and `command-exec` events. Command events are logged through a global `command-after-exec.py` hook, so weekly per-tool counts are collected centrally.
4. Admin endpoints:
   `GET /admin?key=...`
   `GET /api/admin/summary?days=7&key=...`
   `GET /api/admin/weekly-report?weeks=8&format=markdown&key=...`
   `GET /api/admin/weekly-report?weeks=8&format=csv&key=...`
