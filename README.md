# WWPTools (pyRevit Extension)

WWPTools is a pyRevit toolbar extension distributed via GitHub Releases for easy installs and updates.

## For admins (publish updates)
1) Create a GitHub repo named `WWPTools` (public).
2) Publish updates by pushing to `main` (the installer pulls the repo zip from GitHub).
3) If the repo owner or name changes, update it in `installer/WWPTools.wxs` and rebuild the MSI.
4) Attach the MSI to a GitHub release (example: `V1.01`).

## For users (install or update)
1) Download the MSI from the latest GitHub release (example: `V1.01`).
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
