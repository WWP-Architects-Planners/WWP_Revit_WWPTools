# SignPath Setup (GitHub Trusted Build)

This repository includes a manual GitHub Actions workflow:

- `.github/workflows/sign-msi-signpath.yml`

It builds the MSI installers and submits both MSI files to SignPath for signing.

## 1) Required GitHub Secret

- `SIGNPATH_API_TOKEN`

Create this token in SignPath and add it under:
GitHub repo -> Settings -> Secrets and variables -> Actions -> Secrets.

## 2) Required GitHub Variables

Add these under:
GitHub repo -> Settings -> Secrets and variables -> Actions -> Variables.

- `SIGNPATH_ORGANIZATION_ID`
- `SIGNPATH_PROJECT_SLUG`
- `SIGNPATH_SIGNING_POLICY_SLUG`
- `SIGNPATH_ARTIFACT_CONFIG_MAIN`
- `SIGNPATH_ARTIFACT_CONFIG_PACKAGES`

Notes:
- `SIGNPATH_SIGNING_POLICY_SLUG` should match your selected policy (for example `mit` if your policy is named MIT).
- Artifact configuration slugs must match configurations in SignPath for each MSI.

## 3) Run the signing workflow

GitHub repo -> Actions -> `Sign MSI (SignPath)` -> Run workflow

Inputs:
- `version`: release version (example `1.9.1`)
- `sign_packages_msi`: `true` to also sign `WWPTools-Packages-v1.0.0.msi`

## 4) Outputs

The workflow publishes signed artifacts:

- `signed-wwptools-msi`
- `signed-wwptools-packages-msi` (optional)

These can be downloaded from the workflow run page.
