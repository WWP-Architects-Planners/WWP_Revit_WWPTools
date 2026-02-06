### Added
- Update checker on startup with GitHub latest-release comparison
- Local version file `WWPTools.extension/lib/WWPTools.version.json`
- About button icon and version tooltip
- Update WWPTools button (fetch/pull GitHub updates with user confirmation)
- Windows 10-style toggle styling in WPF dialogs

### Changed
- Sheet Scale Updater: sorted sheet list, selectable sheets, target parameter picker (instance + type), and performance tuning
- Copy Color Scheme: source list shows category/area scheme names; target selection is category-only with single overwrite toggle
- App-init toast now reports "outdated" vs "latest" on load

### Fixed
- CPython output flushing errors in Sheet Scale Updater report
- Cancel/exit handling to avoid CPython SystemExit errors
- IronPython update check import error (urllib fallback)
- Update button fallback when Git is missing (opens releases page)
