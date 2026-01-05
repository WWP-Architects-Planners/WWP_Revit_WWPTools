import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit import UI

# pyRevit script tools
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

output = script.get_output()

# ----------------------------------------------------
# Collect Sheets
# ----------------------------------------------------
sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

if not sheets:
    UI.TaskDialog.Show("Sheet Scale Updater", "No sheets found.")
    script.exit()


updated_count = 0
failed_sheets = []
debug_info = []

# ----------------------------------------------------
# Start Transaction
# ----------------------------------------------------
t = Transaction(doc, "Update Sheet Scale")
t.Start()

try:

    for sheet in sheets:
        sheet_name = sheet.Name
        sheet_debug = {"sheet": sheet_name}

        # ----------------------------------------------------
        # Find Titleblock
        # ----------------------------------------------------
        titleblock_instance = None

        for element in FilteredElementCollector(doc, sheet.Id).OfClass(FamilyInstance):
            if element.Category:
                # Revit 2025/2026 compatibility for ElementId comparison
                category_id = element.Category.Id
                titleblock_builtin_id = int(BuiltInCategory.OST_TitleBlocks)
                
                is_titleblock = False
                # Try Revit 2025 method first (IntegerValue property)
                if hasattr(category_id, 'IntegerValue'):
                    is_titleblock = category_id.IntegerValue == titleblock_builtin_id
                else:
                    # Revit 2026 method: convert to long for ElementId constructor
                    try:
                        titleblock_id = ElementId(long(titleblock_builtin_id))
                        is_titleblock = category_id == titleblock_id
                    except:
                        # Fallback: compare as integers
                        is_titleblock = int(category_id) == titleblock_builtin_id
                
                if is_titleblock:
                    titleblock_instance = element
                    break

        if not titleblock_instance:
            sheet_debug["error"] = "No titleblock found"
            failed_sheets.append(sheet_name + " - No titleblock found")
            debug_info.append(sheet_debug)
            continue

        sheet_debug["titleblock_found"] = True

        # ----------------------------------------------------
        # Collect viewport scales
        # ----------------------------------------------------
        viewport_ids = sheet.GetAllViewports()

        if viewport_ids.Count == 0:
            sheet_debug["error"] = "No viewports"
            failed_sheets.append(sheet_name + " - No viewports")
            debug_info.append(sheet_debug)
            continue

        sheet_debug["viewport_count"] = viewport_ids.Count
        scales = set()
        scale_details = []

        for vp_id in viewport_ids:
            viewport = doc.GetElement(vp_id)
            if viewport:

                # Get the view from the viewport to read its scale
                view = doc.GetElement(viewport.ViewId)
                if view and hasattr(view, 'Scale'):
                    sval = view.Scale
                    scale_details.append(sval)
                    if sval > 0:
                        scales.add(sval)

        sheet_debug["scales_read"] = sorted(list(scales))
        sheet_debug["all_viewport_scales"] = scale_details

        if not scales:
            sheet_debug["error"] = "No valid scales found"
            failed_sheets.append(sheet_name + " - No valid scales found")
            debug_info.append(sheet_debug)
            continue

        # determine final scale
        if len(scales) == 1:
            sheet_scale_value = list(scales)[0]
            sheet_debug["scale_selection"] = "single scale"
        else:
            # Multiple scales - set to 0
            sheet_scale_value = 0
            sheet_debug["multiple_scales"] = True
            sheet_debug["scale_selection"] = "multiple scales detected - setting to 0"
            
            # Still check override param for info
            override_param = titleblock_instance.LookupParameter("Show Scale Override")
            sheet_debug["override_param_exists"] = override_param is not None
            
            if override_param:
                sheet_debug["override_value"] = override_param.AsInteger()

        sheet_debug["scale_to_write"] = sheet_scale_value

        # ----------------------------------------------------
        # Write to "Sheet Scale"
        # ----------------------------------------------------
        sheet_scale_param = titleblock_instance.LookupParameter("Sheet Scale")
        sheet_debug["sheet_scale_param_exists"] = sheet_scale_param is not None

        if not sheet_scale_param:
            sheet_debug["error"] = "Sheet Scale parameter missing"
            failed_sheets.append(sheet_name + " - Sheet Scale parameter missing")
            debug_info.append(sheet_debug)
            continue

        sheet_debug["param_storage_type"] = str(sheet_scale_param.StorageType)
        
        try:
            # Handle parameter value setting - Revit 2026 compatibility
            if sheet_scale_param.IsReadOnly:
                sheet_debug["error"] = "Parameter is read-only"
                failed_sheets.append(sheet_name + " - Parameter is read-only")
            elif sheet_scale_param.StorageType == StorageType.Double:
                sheet_scale_param.Set(float(sheet_scale_value))
                sheet_debug["written_as"] = "Double"
                sheet_debug["status"] = "SUCCESS"
                updated_count += 1
            elif sheet_scale_param.StorageType == StorageType.Integer:
                sheet_scale_param.Set(int(sheet_scale_value))
                sheet_debug["written_as"] = "Integer"
                sheet_debug["status"] = "SUCCESS"
                updated_count += 1
            else:
                sheet_scale_param.Set(sheet_scale_value)
                sheet_debug["written_as"] = str(sheet_scale_param.StorageType)
                sheet_debug["status"] = "SUCCESS"
                updated_count += 1

        except Exception as e:
            sheet_debug["error"] = "Failed to set parameter: {}".format(str(e))
            failed_sheets.append(sheet_name + " - Failed to set parameter: {}".format(str(e)))
        
        debug_info.append(sheet_debug)

    t.Commit()

except Exception as e:
    t.RollBack()
    UI.TaskDialog.Show("Error", str(e))
    script.exit()


# ----------------------------------------------------
# Summary Dialog
# ----------------------------------------------------
msg = "Successfully updated {} sheet(s).".format(updated_count)

if failed_sheets:
    msg += "\n\nFailed/Skipped:\n" + "\n".join(failed_sheets[:10])
    if len(failed_sheets) > 10:
        msg += "\n... and {} more".format(len(failed_sheets) - 10)

UI.TaskDialog.Show("Sheet Scale Updater", msg)


# ----------------------------------------------------
# Detailed Debug Output to Console
# ----------------------------------------------------
output.print_md("# Sheet Scale Updater - Detailed Report\n")
output.print_md("**Developed by: Jason Tian**\n\n")
output.print_md("**Updated:** {} sheets\n".format(updated_count))
output.print_md("**Failed/Skipped:** {} sheets\n".format(len(failed_sheets)))

output.print_md("\n## Per-Sheet Details:\n")

for debug in debug_info:
    output.print_md("\n### Sheet: `{}`".format(debug.get("sheet", "Unknown")))
    
    # Status indicator
    if debug.get("status") == "SUCCESS":
        output.print_md("**Status:** [SUCCESS]")
    else:
        output.print_md("**Status:** [FAILED]")
    
    # Titleblock info
    if debug.get("titleblock_found"):
        output.print_md("- **Titleblock:** Found")
    if "error" in debug and "titleblock" in debug["error"].lower():
        output.print_md("- **Titleblock:** NOT FOUND [WARNING]")
    
    # Viewport info
    if "viewport_count" in debug:
        output.print_md("- **Viewports:** {} found".format(debug["viewport_count"]))
    
    # Scale read information
    if "all_viewport_scales" in debug:
        output.print_md("- **All Viewport Scales (read):** {}".format(debug["all_viewport_scales"]))
    
    if "scales_read" in debug:
        output.print_md("- **Valid Scales (after filter):** {}".format(debug["scales_read"]))
    
    # Scale selection method
    if "scale_selection" in debug:
        output.print_md("- **Scale Selection Method:** {}".format(debug["scale_selection"]))
    
    # Multiple scales debug info
    if debug.get("multiple_scales"):
        output.print_md("- **Multiple Scales Detected:** Yes")
        if "override_param_exists" in debug:
            output.print_md("  - **Override Parameter Exists:** {}".format(debug["override_param_exists"]))
        if "override_value" in debug:
            output.print_md("  - **Override Value:** {}".format(debug["override_value"]))
    
    # Scale to write
    if "scale_to_write" in debug:
        output.print_md("- **Scale to Write:** {}".format(debug["scale_to_write"]))
    
    # Parameter info
    if "sheet_scale_param_exists" in debug:
        output.print_md("- **Sheet Scale Parameter Exists:** {}".format(debug["sheet_scale_param_exists"]))
    
    if "param_storage_type" in debug:
        output.print_md("- **Parameter Storage Type:** {}".format(debug["param_storage_type"]))
    
    if "written_as" in debug:
        output.print_md("- **Written As:** {}".format(debug["written_as"]))
    
    # Error details
    if "error" in debug:
        output.print_md("- **Error:** {}".format(debug["error"]))
