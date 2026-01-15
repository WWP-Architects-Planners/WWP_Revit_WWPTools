# 1. Project Setup

This page documents setup steps and tools found under the WWPTools/SvNTool add-on tab.

## Prerequisites and setup

### Install pyRevit (required)
Location: Revit installer
Purpose: Required host for WWPTools and SvNTool extensions.
How to use:
1. Run the installer from `N:\Library\Design Software\Autodesk\Revit\! Add-Ins\pyRevit\pyRevit_4.8.12.22247_signed.exe`.
2. Accept the agreement and continue through all prompts.
3. Open Revit and click "Always Load" on the security prompt.

### Install WWPTools MSI
Location: GitHub Releases
Purpose: Install the WWPTools add-in via the official MSI.
How to use:
1. Download the latest MSI from `https://github.com/SvN-Architects-Planners/WWPTools/releases/latest`.
   - Current release: `WWPTools-v1.1.3.msi`
2. Run the MSI and follow the prompts.
3. Open Revit and verify the WWPTools tab is visible.

### Register SvNTool extensions in pyRevit (legacy)
Location: Revit > pyRevit tab > pyRevit panel > Settings
Purpose: Add the SvNTool extension folder so the tab appears in Revit.
How to use:
1. In pyRevit Settings, under Custom Extension Directories, click "Add folder".
2. Add `N:\Library\Design Software\Autodesk\Revit\Dynamo\SvN Tool`.
3. Click "Save Settings and Reload".

### Dynamo first-time setup (legacy)
Location: Revit > Manage tab > Dynamo
Purpose: Accept the Dynamo license and install core components.
How to use:
1. Launch Dynamo from Revit and accept the terms.
2. When prompted, choose "Install on C drive".

### Install SvNTool Dynamo packages (legacy)
Location: SvNTool tab > Install Packages
Purpose: Install required Dynamo packages for all tools.
How to use:
1. In SvNTool, use Install Packages and select "Install on C drive".
2. In Dynamo, go to Preferences and open the Package Manager settings.
3. Add package path `C:\dynpackages\All Version`.
4. Close Dynamo to apply settings.

### Optional: enable EF_Tools extension (legacy)
Location: Revit > pyRevit tab > Extensions
Purpose: Enable EF_Tools add-on features.
How to use:
1. Select "EF-Tools" and click "Enable Extension".
2. If prompted, install to the C drive.
3. In Dynamo Preferences, set Default Python Engine to IronPython2.
4. Verify packages are installed:
   - Archi-lab.net
   - bimorphNodes
   - Clockwork for Dynamo 2.x
   - Crumple
   - Data-Shapes
   - DynamoIronPython 2.7
   - Elk
   - Genius Loci
   - Rhythm
   - spring nodes
   - SvN Packages

## Tool reference

### Line Type Tool
Location: WWPTools > 1. Project Setup > Line Type Tool
Purpose: Import and standardize line types in the project.
Source file: `N:\Library\Design Software\Autodesk\Revit\Standards\Line Type\Linetypes.xlsx`
How to use:
1. Run the tool and follow prompts.
2. Review the results and undo if needed.

### Add Project Parameter (Pulldown)
Purpose: Import standard shared parameters or build a template.

#### Create Template
Location: WWPTools > 1. Project Setup > Add Project Parameter > Create Template
Purpose: Create an Excel template with shared parameters listed in the Variables sheet.
How to use:
1. Have the target sheet(s) available or selected as prompted.
2. Run the tool and follow prompts.
3. Review the results and undo if needed.

#### Import from Excel
Location: WWPTools > 1. Project Setup > Add Project Parameter > Import from Excel
Purpose: Import shared parameters from a workbook.
Source file: `N:\Library\Design Software\Autodesk\Revit\Shared Parameters\Shared Parameters Import.xlsx`
How to use:
1. Prepare the Excel file you want to import.
2. Run the tool and follow prompts.
3. Review the results and undo if needed.

### Levels Setup (Pulldown)
Purpose: Manage levels and generate views.

#### Levels Setup
Location: WWPTools > 1. Project Setup > Levels Setup > Levels Setup
Purpose: Add or remove levels based on a target floor count.
How to use:
1. Enter the desired number of floors.
2. Confirm the list of levels being added or deleted.

#### Views Creator
Location: WWPTools > 1. Project Setup > Levels Setup > Views Creator
Purpose: Create Area Plans, Floor Plans, and RCPs for selected levels.
How to use:
1. Choose categories and view types (Area Plan, Floor Plan, or RCP).
2. Select the floors to create.
3. Click "Set Value" and wait for results.

### Project Upgrader
Location: SvNTool tab > Project Upgrader
Purpose: Batch-upgrade families or models to the current Revit version.
How to use:
1. Select the source folder and click Select Folder.
2. Wait for the tool to process and rename upgraded files.

### Copy Color Scheme
Location: SvNTool tab > Copy Color Scheme
Purpose: Copy color scheme settings between schemes.
How to use:
1. Select the source and target color scheme.
2. Run the tool to copy the settings.

### Copy Filters
Location: SvNTool tab > Copy Filters
Purpose: Copy filter overrides between view templates and views.
How to use:
1. Choose one of the actions:
   - Copy Current Filters to Multiple Templates.
   - Copy Filters to Current View.
   - Copy Filters to Template.
2. Follow the prompts.

### Copy Parameter
Location: SvNTool tab > Copy Parameter
Purpose: Copy values from one instance parameter to another with optional find/replace.
How to use:
1. Select elements to update.
2. Enter source and target parameter names.
3. (Optional) Enter find/replace strings.
4. Click "Set Values".

### CAD Line Tool (Current Selection)
Location: SvNTool tab > CAD Line Tool
Purpose: Convert linked CAD layers into detail line groups in the current view.
How to use:
1. Select the linked CAD file in the view.
2. Run the tool to generate grouped detail lines by layer.

### Door Tool
Location: SvNTool tab > Door Tool
Purpose: Bulk edit door data and types.
How to use:
1. Choose the action:
   - Publish Fire Rating to Doors: copy wall FRR to door parameter (FRR Walls or FRR).
   - Select Doors in View: selects all doors in the view.
   - Door Type Duplicator: duplicates a door type with dimensions.
   - Get Door Number from Room: writes door numbers from To Room values.
2. Follow prompts.
Note: Get Door Number from Room overwrites all doors in the project and does not work for model group doors.

### Replace Type Name
Location: SvNTool tab > Replace Type Name
Purpose: Batch rename type names.
How to use:
1. Enter the search text and replacement text.
2. Run the tool.

### Fire Rating Tool
Location: SvNTool tab > Fire Rating Tool
Purpose: Create or clear fire rating detail lines.
How to use:
1. Choose an action:
   - Create Non-Rated Line for all walls in current view.
   - Create Fire Rating Lines for all walls in current view.
   - Create Fire Rating Lines for all walls in views with FRR indicator.
   - Create Fire Rating Lines for selected walls in current view.
   - Clear all fire rating lines in current view.
2. Ensure FRR values match expected ratings: `0HR`, `.75HR`, `1HR`, `1.5HR`, `2HR`, `3HR`.
3. Optional: Convert Lines to Detail Items (requires family `N:\Library\Design Software\Autodesk\Revit\Families\Annotations\Families\SvN_FireRating_FRR Line Based.rfa`).

### Match Filled Region
Location: SvNTool tab > Match Filled Region
Purpose: Standardize filled region graphics across project and families.
How to use:
1. Run the tool to match all filled regions to project standards.

### Parking Count in Room
Location: SvNTool tab > Parking Count in Room
Purpose: Count parking by room and fill the parking occupancy room tag.
How to use:
1. Fill required parameters on parking families.
2. Select a room and run the tool.
3. The tag `_SvN_A_RoomTag_ParkingOccupancy` is populated.

### Project Stat Manager
Location: SvNTool tab > Project Setup / Update Stats
Purpose: Export schedules to CSV and manage project statistics.
How to use:
1. Copy sample Excel files from `N:\Library\Design Software\Autodesk\Revit\Sample Files\stats`.
2. Rename files with the project number.
3. Run `projectsetup`, select schedules (SvN_*), and choose the export folder:
   `project\6 WORK_INFO\6-2 Stats\Queries`.
4. Click Export.
5. To update later, run `upgradestats`.
6. In Excel, load each CSV into the corresponding SvN sheet (Data > From Text/CSV).

### Marketing View Maker
Location: SvNTool tab > Marketing View Maker
Purpose: Create marketing blackline views and sheets for suites.
How to use:
1. Select an area and its entry door.
2. Choose Make Single Suite or Make Multiple KeyPlans.
3. Pick view templates and titleblock.
Notes:
- Creates backup sheets with `- bak` suffix if sheets already exist.
- Auto-crops and rotates to align the door direction.

### AutoFill SvN_Area
Location: SvNTool tab > AutoFill SvN_Area
Purpose: Auto-fill SvN area parameter for deductions and GFA counts.
Note: This tool is being deprecated due to the key area schedule method.

### Context Builder (OSM)
Location: SvNTool tab > Context Builder (OSM)
Purpose: Build site context from OpenStreetMap or CAD.
How to use:
1. Export an OSM file from OpenStreetMap.
2. Create a new Conceptual Mass family in Revit.
3. Use one of the tools:
   - Building Importer: select OSM file and units (mm or meters).
   - CAD Builder: select CAD file and layer; keep defaults unless needed.
   - Road Importer: select OSM file and units; keep default widths unless needed.
4. Load into project; avoid "Copy to In-Place-Mass" unless necessary.

### Random Tree
Location: SvNTool tab > Random Tree
Purpose: Randomize tree family rotation and scale.
How to use:
1. Run the tool to randomize tree direction and size within 10 percent.

### Sync Mass Tool
Location: SvNTool tab > Sync Mass Tool
Purpose: Push mass data to mass floors for GCA/GFA and FSI reporting.
How to use:
1. Create a mass and mass levels.
2. Enter data per mass.
3. Run Sync Mass Tool to transfer values to mass floors.

### Purge Views/Sheets
Location: SvNTool tab > Purge Views/Sheets
Purpose: Clean up views and sheets.
How to use:
1. Choose Clean Selected Sheet sets and Views or Delete Unused Views.
2. Select the print set if prompted.

### Duplicate Views
Location: SvNTool tab > Duplicate Views
Purpose: Bulk duplicate views and update names.
How to use:
1. Enter the text to replace (or leave blank).
2. Enter the new text or suffix.
3. Choose Duplicate View option.

### Schedule Exporter
Location: SvNTool tab > Schedule Exporter
Purpose: Export schedules to Excel and rename views/sheets in bulk.
How to use:
1. Select schedules to export and choose the Excel host file.
2. Provide a name and run the export.
3. Use Replace View Name to batch rename views (supports prefix and suffix).

### Sheets Manager
Location: SvNTool tab > Sheets Manager
Purpose: Duplicate or delete sheets and views.
How to use:
1. Select the sheet(s).
2. Choose whether to duplicate views on the sheet.
3. Use the action:
   - Delete Selected Sheet and its views.
   - Delete Views from Current Sheet.
   - Sheet Duplicator.
Note: Open sheets will not be deleted.

### Project Cleaner
Location: SvNTool tab > Project Cleaner
Purpose: Clean local ACC and temp files.
How to use:
1. Close and save all active files.
2. Open a blank dummy file.
3. Run Project Cleaner.

### Wipe Schema
Location: SvNTool tab > Wipe Schema
Purpose: Clean extra Revit schema data.
How to use:
1. Run the tool to remove stale schema data.
