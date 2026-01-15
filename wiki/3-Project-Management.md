# 3. Project Management

This page documents the tools found under this WWPTools panel.

## Copy Color Scheme
Location: WWPTools > 3. Project Management > Copy Color Scheme
Screenshot: (Add later)
Purpose: Copy color scheme settings between schemes.
How to use:
1. Select the source and target color scheme.
2. Run the tool to copy the settings.

## Copy Parameter
Location: WWPTools > 3. Project Management > Copy Parameter
Screenshot: (Add later)
Purpose: Copy values from one instance parameter to another with optional find/replace.
How to use:
1. Select elements to update.
2. Enter source and target parameter names.
3. (Optional) Enter find/replace strings.
4. Click "Set Values".

## Parking Count in Room
Location: WWPTools > 3. Project Management > Parking Count in Room
Screenshot: (Add later)
Purpose: Count parking by room and fill the parking occupancy room tag.
How to use:
1. Fill required parameters on parking families.
2. Select a room and run the tool.
3. The tag `_SvN_A_RoomTag_ParkingOccupancy` is populated.

## Replace Type Name
Location: WWPTools > 3. Project Management > Replace Type Name
Screenshot: (Add later)
Purpose: Batch rename type names.
How to use:
1. Enter the search text and replacement text.
2. Run the tool.

## Match Filled Region
Location: WWPTools > 3. Project Management > Match Filled Region
Screenshot: (Add later)
Purpose: Standardize filled region graphics across project and families.
How to use:
1. Run the tool to match all filled regions to project standards.

## Sheet Scale Updater
Location: WWPTools > 3. Project Management > Sheet Scale Updater
Screenshot: (Add later)
Purpose: Update titleblock Sheet Scale parameter from sheet views.
How to use:
1. Have the target sheet(s) available or selected as prompted.
2. Run the tool and follow prompts.

## AutoFill WWP_Area (Pulldown)
Purpose: Group of related tools.
Note: This tool is being deprecated due to the key area schedule method.

### AutoFill ALL WWP_Area
Location: WWPTools > 3. Project Management > AutoFill WWP_Area > AutoFill ALL WWP_Area
Screenshot: (Add later)
Purpose: Download all WWP standard data based on Unit Type.
How to use:
1. Run the tool and follow prompts.

### AutoFill Selected WWP_Area
Location: WWPTools > 3. Project Management > AutoFill WWP_Area > AutoFill Selected WWP_Area
Screenshot: (Add later)
Purpose: Download selected WWP standard data based on Unit Type.
How to use:
1. Select the target elements before running the tool.
2. Run the tool and follow prompts.

## Copy Filters (Pulldown)
Purpose: Copy filter graphic overrides between view templates and views.

### Copy Current Filters to Multiple Templates
Location: WWPTools > 3. Project Management > Copy Filters > Copy Current Filters to Multiple Templates
Screenshot: (Add later)
Purpose: Copy filter graphic overrides from current view to multiple templates.
How to use:
1. Open the target view you want to affect.
2. Run the tool and follow prompts.

### Copy Filters to Current View
Location: WWPTools > 3. Project Management > Copy Filters > Copy Filters to Current View
Screenshot: (Add later)
Purpose: Copy filter graphic overrides from one template to the current view.
How to use:
1. Open the target view you want to affect.
2. Run the tool and follow prompts.

### Copy Filters to Template
Location: WWPTools > 3. Project Management > Copy Filters > Copy Filters to Template
Screenshot: (Add later)
Purpose: Copy filter graphic overrides from one template to another.
How to use:
1. Run the tool and follow prompts.

## Door Tool (Pulldown)
Purpose: Bulk edit door data and types.

### Door Type Duplicator
Location: WWPTools > 3. Project Management > Door Tool > Door Type Duplicator
Screenshot: (Add later)
Purpose: Duplicate a selected door type with specified dimensions.
How to use:
1. Make sure door elements are available in the active view or selection.
2. Run the tool and follow prompts.

### Get Door Number from Room
Location: WWPTools > 3. Project Management > Door Tool > Get Door Number from Room
Screenshot: (Add later)
Purpose: Write door numbers from To Room values into the door Mark.
How to use:
1. Make sure door elements are available in the active view or selection.
2. Run the tool and follow prompts.
Note: This overwrites all doors in the project and does not work for model group doors.

### Publish Fire Rating to Doors
Location: WWPTools > 3. Project Management > Door Tool > Publish Fire Rating to Doors
Screenshot: (Add later)
Purpose: Copy wall fire rating (FRR) to a door parameter.
How to use:
1. Select the target doors and walls.
2. Enter the wall parameter name (FRR Walls or FRR).
3. Run the tool and follow prompts.

### Select Doors in View
Location: WWPTools > 3. Project Management > Door Tool > Select Doors in View
Screenshot: (Add later)
Purpose: Select all doors in the current view.
How to use:
1. Open the target view you want to affect.
2. Run the tool and follow prompts.

## Fire Rating Tool (Pulldown)
Purpose: Create or clear fire rating detail lines for walls.

### Clear All fire rating lines in Current View
Location: WWPTools > 3. Project Management > Fire Rating Tool > Clear All fire rating lines in Current View
Screenshot: (Add later)
Purpose: Clear all auto-generated fire rating lines in the current view.
How to use:
1. Open the target view you want to affect.
2. Run the tool and follow prompts.

### Convert Lines to Detail Items
Location: WWPTools > 3. Project Management > Fire Rating Tool > Convert Lines to Detail Items
Screenshot: (Add later)
Purpose: Convert rated detail lines into line-based detail components.
How to use:
1. Load `N:\Library\Design Software\Autodesk\Revit\Families\Annotations\Families\SvN_FireRating_FRR Line Based.rfa`.
2. Run the tool and follow prompts.

### Create Fire Rating Lines for ALL walls in Current View
Location: WWPTools > 3. Project Management > Fire Rating Tool > Create Fire Rating Lines for ALL walls in Current View
Screenshot: (Add later)
Purpose: Create rated lines using wall FRR values.
How to use:
1. Open the target view you want to affect.
2. Run the tool and follow prompts.
Notes:
- Expected FRR values: `0HR`, `.75HR`, `1HR`, `1.5HR`, `2HR`, `3HR`.

### Create Fire Rating Lines for All walls in views with FRR indicator
Location: WWPTools > 3. Project Management > Fire Rating Tool > Create Fire Rating Lines for All walls in views with FRR indicator
Screenshot: (Add later)
Purpose: Create fire rating lines for all views that end with "FRR".
How to use:
1. Run the tool and follow prompts.

### Create Fire Rating Lines for Selected walls in Current View
Location: WWPTools > 3. Project Management > Fire Rating Tool > Create Fire Rating Lines for Selected walls in Current View
Screenshot: (Add later)
Purpose: Create rated lines for selected walls.
How to use:
1. Select walls in the current view.
2. Run the tool and follow prompts.

### Create Non-Rated Line for ALL walls in Current View
Location: WWPTools > 3. Project Management > Fire Rating Tool > Create Non-Rated Line for ALL walls in Current View
Screenshot: (Add later)
Purpose: Create solid centerline detail lines for all walls regardless of FRR values.
How to use:
1. Open the target view you want to affect.
2. Run the tool and follow prompts.

## Marketing View Maker (Pulldown)
Purpose: Create marketing blackline views and sheets (Revit 2023+).

### Make MULTIPLE KeyPlans from Selected Area
Location: WWPTools > 3. Project Management > Marketing View Maker > Make MULTIPLE KeyPlans from Selected Area
Screenshot: (Add later)
Purpose: Create multiple keyplans based on selected areas.
How to use:
1. Select the target areas.
2. Run the tool and follow prompts.

### Make Single Suite (Recommended)
Location: WWPTools > 3. Project Management > Marketing View Maker > Make Single Suite (Recommended)
Screenshot: (Add later)
Purpose: Create suite blackline views and sheets from a selected area and entry door.
How to use:
1. Select an area and its entry door.
2. Choose view templates and titleblock.
3. Run the tool and follow prompts.
Notes:
- Creates backup sheets with `- bak` suffix if sheets already exist.
- Auto-crops and rotates to align the door direction.

## Project Stat Manager (Pulldown)
Purpose: Export schedules to CSV and manage project statistics.

### Projectsetup
Location: WWPTools > 3. Project Management > Project Stat Manager > Projectsetup
Screenshot: (Add later)
Purpose: Export selected schedules (SvN_*) to CSV.
How to use:
1. Copy sample Excel files from `N:\Library\Design Software\Autodesk\Revit\Sample Files\stats`.
2. Rename files with the project number.
3. Choose the export folder `project\6 WORK_INFO\6-2 Stats\Queries`.
4. Run Projectsetup and click Export.

### UpdateStats
Location: WWPTools > 3. Project Management > Project Stat Manager > UpdateStats
Screenshot: (Add later)
Purpose: Update all CSV exports for the stats workbook.
How to use:
1. Run UpdateStats to refresh CSV outputs.
2. In Excel, load each CSV into the corresponding SvN sheet (Data > From Text/CSV).
