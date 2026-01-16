# Changelog

All notable changes to WWPTools will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.3] - 2026-01-09

### Fixed
- Installer License Agreement dialog UI errors (first active control and tab order)
- MSI build step now patches dialog metadata after banner removal

## [1.1.4] - 2026-01-16

### Added
- CPython versions of the Sheet Manager tools (delete sheets, delete sheet views, duplicate sheets)
- Sheet Duplicator UI with duplicate options, prefix/suffix, and view duplication controls
- Replace View Name CPython tool with find/replace/prefix/suffix inputs
- Delete Unused Views CPython tool with confirmation prompt

### Changed
- Renumber Sheets now uses a single dialog for selection and starting number
- Sheet Manager Dynamo graphs moved into archive folders

### Fixed
- UI helper now handles null/non-iterable inputs for sheet renumber selection

## [1.1.2] - 2026-01-09

### Changed
- Updated Mass Context tool bundles for Random Plants, Mass ID Tool, Detail Line CAD, and Rename Materials
- Updated Randomtree and CADLine scripts
- Updated `WWPTools.tab/bundle.yaml`

## [1.1.1] - 2026-01-08

### Changed
- Multiple Schedules Exporter completely rewritten for CPython with WinForms UI
- Excel export now uses openpyxl via CSV pipeline; CSV export unchanged
- Export tool now remembers last schedule selection, last Excel file location, and last CSV folder
- Schedule list excludes legend views
- Sync task excludes archive folders
- Refactored Export2Ex tool from pulldown menu to single pushbutton
- Consolidated multiple schedule export tools into unified interface
- Improved user experience with streamlined export workflow

### Removed
- Deprecated separate SingleSchedule and MultipleSchedules buttons
- Removed archive Dynamo scripts for schedule export

## [1.1.0] - 2026-01-07

### Added
- Reorganized Add Project Parameter tool into a pulldown menu with multiple options
  - Create Template functionality
  - Import from Excel functionality

### Changed
- Converted CopyParameter tool from Dynamo to Python script for better performance
- Refactored project structure for better organization
- Updated various tool scripts and configurations
- Updated extension hooks and library utilities
- Enhanced bundle configurations across multiple tools
- Improved line ending consistency across files

### Removed
- Deprecated package copy tools (copyallpackages pulldown)

### Archived
- Old implementations moved to archive folders for reference

## [1.0.0] - 2026-01-06

### Added
- Initial release of WWPTools
- Published with package for WW+P users
- Useful tools and shortcuts for better productivity
- Mass Context tools
- Project Management tools
- Project Setup tools
- Revit Cleanup tools
- Views Sheet Manager tools

[1.1.4]: https://github.com/jason-svn/WWPTools/releases/tag/V1.1.4
[1.1.3]: https://github.com/jason-svn/WWPTools/releases/tag/V1.1.3
[1.1.2]: https://github.com/jason-svn/WWPTools/releases/tag/V1.1.2
[1.1.1]: https://github.com/jason-svn/WWPTools/releases/tag/V1.1.1
[1.1.0]: https://github.com/jason-svn/WWPTools/releases/tag/V1.1.0
[1.0.0]: https://github.com/jason-svn/WWPTools/releases/tag/V1.0.0
