# Changelog

All notable changes to WWPTools will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.1] - 2026-01-08

### Changed
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

[1.1.0]: https://github.com/jason-svn/WWPTools/releases/tag/V1.1.0
[1.0.0]: https://github.com/jason-svn/WWPTools/releases/tag/V1.0.0
