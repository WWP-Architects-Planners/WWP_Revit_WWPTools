# WWPTools (pyRevit Extension)

WWPTools is a pyRevit toolbar extension distributed via GitHub Releases for easy installs and updates.

## For admins (publish updates)
1) Create a GitHub repo named `WWPTools` (public).
2) Upload a release asset zip that contains the folder `pyRevitTool.extension` at the root of the zip.
3) Publish a new Release for each update (the installer pulls the latest release).

## For users (install or update)
1) Download `Install_WWPTools.bat` from the repo.
2) Double-click it. It installs/updates the extension in:
   `%APPDATA%\pyRevit\Extensions`

## Dynamo packages (optional)
If you want shared Dynamo packages without overwriting user packages:
1) Download `Install_DynamoPackages.bat` from the repo.
2) Double-click it. It installs packages to `C:\dynpackages` and adds that
   folder to Dynamo's Custom Package Folders (for any installed Dynamo Revit versions).

## Customize the repo owner
Edit `Update_WWPTools.ps1` and set `RepoOwner` to your GitHub org/user name.
