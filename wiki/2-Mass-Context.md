# 2. Mass Context

This page documents the tools found under this WWPTools panel.

## Bulk Rename Materials
Location: WWPTools > 2. Mass Context > Bulk Rename Materials
Screenshot: (Add later)
Purpose: Find/replace text across all material names in the current model.
How to use:
1. Run the tool and follow prompts.
2. Review the results and undo if needed.

## CAD Line Tool
Location: WWPTools > 2. Mass Context > CAD Line Tool
Screenshot: (Add later)
Purpose: Convert linked CAD layers into detail line groups in the current view.
How to use:
1. Select the linked CAD file in the view.
2. Run the tool to generate grouped detail lines by layer.

## Random Tree
Location: WWPTools > 2. Mass Context > Random Tree
Screenshot: (Add later)
Purpose: Randomize tree family rotation and scale within a small range.
How to use:
1. Run the tool to randomize tree direction and size within 10 percent.

## Sync Mass Tool
Location: WWPTools > 2. Mass Context > Sync Mass Tool
Screenshot: (Add later)
Purpose: Push mass data to mass floors for GCA/GFA and FSI reporting.
How to use:
1. Create a mass and mass levels.
2. Enter data per mass.
3. Run Sync Mass Tool to transfer values to mass floors.

## Context Builder (Pulldown)
Purpose: Create context building and site elements based on OSM or CAD.

### Building Importer
Location: WWPTools > 2. Mass Context > Context Builder > Building Importer
Screenshot: (Add later)
Purpose: Create context buildings from an OpenStreetMap (OSM) file.
How to use:
1. Export an OSM file from OpenStreetMap.
2. Create a new Conceptual Mass family in Revit.
3. Select the OSM file and units (mm or meters).
4. Load into the project when complete.

### CADBuilder
Location: WWPTools > 2. Mass Context > Context Builder > CADBuilder
Screenshot: (Add later)
Purpose: Create context buildings from a CAD file.
How to use:
1. Select the CAD file and the target layer.
2. Keep default settings unless the project needs changes.
3. Run CADBuilder and load into the project when complete.

### CADBuilder2023
Location: WWPTools > 2. Mass Context > Context Builder > CADBuilder2023
Screenshot: (Add later)
Purpose: CADBuilder workflow for Revit 2023.
How to use:
1. Select the CAD file and the target layer.
2. Keep default settings unless the project needs changes.
3. Run CADBuilder2023 and load into the project when complete.

### Roads Importer
Location: WWPTools > 2. Mass Context > Context Builder > Roads Importer
Screenshot: (Add later)
Purpose: Create roads and parks from an OSM file.
How to use:
1. Select the OSM file and units (mm or meters).
2. Keep default widths if you are not sure.
3. Load into the project when complete.

Notes:
- Avoid "Copy to In-Place-Mass" unless you need editable mass in the model, as it can slow down the project.
