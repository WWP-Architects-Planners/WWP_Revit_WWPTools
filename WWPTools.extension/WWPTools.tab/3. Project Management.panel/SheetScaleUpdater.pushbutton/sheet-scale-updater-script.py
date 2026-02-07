#!python3
import clr
import sys
clr.AddReference('RevitAPI')
from Autodesk.Revit import DB, UI

# pyRevit script tools
from pyrevit import script
from WWP_uiUtils import uiUtils_select_indices, uiUtils_select_items_with_mode, uiUtils_alert

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

output = script.get_output()
config = script.get_config()

def _print_md(text):
    try:
        if output:
            output.print_md(text)
            return
    except Exception:
        pass
    try:
        print(text)
        if hasattr(sys.stdout, "flush"):
            sys.stdout.flush()
    except Exception:
        pass

# wait dialog removed (lazy loading filter values)

def _is_titleblock(element):
    if not element or not element.Category:
        return False
    category_id = element.Category.Id
    titleblock_builtin_id = int(DB.BuiltInCategory.OST_TitleBlocks)
    # Try Revit 2025 method first (IntegerValue property)
    if hasattr(category_id, 'IntegerValue'):
        return category_id.IntegerValue == titleblock_builtin_id
    # Revit 2026 method: convert to long for ElementId constructor
    try:
        titleblock_id = DB.ElementId(int(titleblock_builtin_id))
        return category_id == titleblock_id
    except:
        # Fallback: compare as integers
        try:
            return int(category_id) == titleblock_builtin_id
        except:
            return False

def _collect_titleblocks():
    return list(
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
        .WhereElementIsNotElementType()
        .ToElements()
    )

def _get_param_entries(titleblock_instances):
    entries = {}
    for tb in titleblock_instances:
        if not tb:
            continue
        try:
            for p in tb.Parameters:
                if not p or not p.Definition:
                    continue
                name = p.Definition.Name
                if name:
                    key = "{} [Instance]".format(name)
                    entries[key] = {"name": name, "scope": "instance"}
        except Exception:
            continue
        try:
            symbol = tb.Symbol if hasattr(tb, "Symbol") else None
            if symbol:
                for p in symbol.Parameters:
                    if not p or not p.Definition:
                        continue
                    name = p.Definition.Name
                    if name:
                        key = "{} [Type]".format(name)
                        entries[key] = {"name": name, "scope": "type"}
        except Exception:
            continue
    keys = sorted(entries.keys())
    return keys, entries

def main():
    sheets = list(DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements())
    sheets.sort(key=lambda s: (s.SheetNumber or "", s.Name or ""))
    if not sheets:
        UI.TaskDialog.Show("Sheet Scale Updater", "No sheets found.")
        return

    updated_count = 0
    failed_sheets = []
    debug_info = []

    sheet_items = []
    sheet_by_index = []
    for sheet in sheets:
        number = sheet.SheetNumber or ""
        name = sheet.Name or ""
        sheet_items.append("{} - {}".format(number, name))
        sheet_by_index.append(sheet)

    last_sheet_ids = getattr(config, "sheet_ids", []) or []
    last_param_name = getattr(config, "sheet_scale_param_name", "") or ""
    last_param_scope = getattr(config, "sheet_scale_param_scope", "") or ""
    prechecked_indices = []
    if last_sheet_ids:
        for i, s in enumerate(sheet_by_index):
            if s.Id.IntegerValue in last_sheet_ids:
                prechecked_indices.append(i)

    try:
        selected_indices, _ = uiUtils_select_items_with_mode(
            sheet_items,
            title="Sheet Scale Updater",
            prompt="Select sheets to update:",
            mode_labels=("", ""),
            default_mode=0,
            prechecked_indices=prechecked_indices,
            width=980,
            height=620,
        )
    except Exception as ex:
        UI.TaskDialog.Show("Sheet Scale Updater", "WPF UI error:\n{}".format(str(ex)))
        return

    if not selected_indices:
        uiUtils_alert("No sheets selected. Operation cancelled.", "Sheet Scale Updater")
        return

    selected_sheets = [sheet_by_index[i] for i in selected_indices]

    titleblocks = _collect_titleblocks()
    if not titleblocks:
        UI.TaskDialog.Show("Sheet Scale Updater", "No titleblock found on any sheet.")
        return

    selected_titleblocks = []
    for s in selected_sheets:
        tb = None
        for t in titleblocks:
            if t.OwnerViewId == s.Id:
                tb = t
                break
        if tb:
            selected_titleblocks.append(tb)

    param_labels, param_entries = _get_param_entries(selected_titleblocks or titleblocks)
    if not param_labels:
        UI.TaskDialog.Show(
            "Sheet Scale Updater",
            "No parameters found on selected titleblock instances.",
        )
        return

    target_param_name = last_param_name
    target_param_scope = last_param_scope
    target_label = None
    if target_param_name and target_param_scope:
        target_label = "{} [{}]".format(target_param_name, "Instance" if target_param_scope == "instance" else "Type")
    if not target_label or target_label not in param_entries:
        try:
            selected = uiUtils_select_indices(
                param_labels,
                title="Sheet Scale Updater",
                prompt="Select titleblock parameter to write scale to:",
                multiselect=False,
                width=720,
                height=560,
            )
        except Exception as ex:
            UI.TaskDialog.Show("Sheet Scale Updater", "WPF UI error:\n{}".format(str(ex)))
            return

        if not selected:
            uiUtils_alert("No parameter selected. Operation cancelled.", "Sheet Scale Updater")
            return

        target_label = param_labels[selected[0]]
        entry = param_entries.get(target_label)
        target_param_name = entry["name"] if entry else target_label
        target_param_scope = entry["scope"] if entry else "instance"

    config.sheet_ids = [s.Id.IntegerValue for s in selected_sheets]
    config.sheet_scale_param_name = target_param_name
    config.sheet_scale_param_scope = target_param_scope
    script.save_config()

    titleblocks_by_sheet = {}
    for tb in titleblocks:
        try:
            titleblocks_by_sheet[tb.OwnerViewId] = tb
        except Exception:
            continue

    t = DB.Transaction(doc, "Update Sheet Scale")
    t.Start()
    try:
        view_scale_cache = {}
        for sheet in selected_sheets:
            sheet_name = sheet.Name
            sheet_debug = {"sheet": sheet_name}
            print("Processing sheet: {}".format(sheet_name))

            titleblock_instance = titleblocks_by_sheet.get(sheet.Id)
            if not titleblock_instance:
                sheet_debug["error"] = "No titleblock found"
                failed_sheets.append(sheet_name + " - No titleblock found")
                debug_info.append(sheet_debug)
                continue

            sheet_debug["titleblock_found"] = True
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
                if not viewport:
                    continue
                view_id = viewport.ViewId
                if view_id in view_scale_cache:
                    sval = view_scale_cache[view_id]
                else:
                    view = doc.GetElement(view_id)
                    sval = view.Scale if view and hasattr(view, 'Scale') else None
                    view_scale_cache[view_id] = sval
                if sval is not None:
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

            if len(scales) == 1:
                sheet_scale_value = list(scales)[0]
                sheet_debug["scale_selection"] = "single scale"
            else:
                sheet_scale_value = 0
                sheet_debug["multiple_scales"] = True
                sheet_debug["scale_selection"] = "multiple scales detected - setting to 0"
                override_param = titleblock_instance.LookupParameter("Show Scale Override")
                sheet_debug["override_param_exists"] = override_param is not None
                if override_param:
                    sheet_debug["override_value"] = override_param.AsInteger()

            sheet_debug["scale_to_write"] = sheet_scale_value

            if target_param_scope == "type":
                symbol = titleblock_instance.Symbol if hasattr(titleblock_instance, "Symbol") else None
                sheet_scale_param = symbol.LookupParameter(target_param_name) if symbol else None
            else:
                sheet_scale_param = titleblock_instance.LookupParameter(target_param_name)
            sheet_debug["target_param_name"] = target_param_name
            sheet_debug["target_param_scope"] = target_param_scope
            sheet_debug["sheet_scale_param_exists"] = sheet_scale_param is not None

            if not sheet_scale_param:
                sheet_debug["error"] = "Target parameter missing"
                failed_sheets.append(sheet_name + " - Target parameter missing: {}".format(target_param_name))
                debug_info.append(sheet_debug)
                continue

            sheet_debug["param_storage_type"] = str(sheet_scale_param.StorageType)
            try:
                if sheet_scale_param.IsReadOnly:
                    sheet_debug["error"] = "Parameter is read-only"
                    failed_sheets.append(sheet_name + " - Parameter is read-only")
                elif sheet_scale_param.StorageType == DB.StorageType.Double:
                    sheet_scale_param.Set(float(sheet_scale_value))
                    sheet_debug["written_as"] = "Double"
                    sheet_debug["status"] = "SUCCESS"
                    updated_count += 1
                elif sheet_scale_param.StorageType == DB.StorageType.Integer:
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
        return

    msg = "Successfully updated {} sheet(s).".format(updated_count)
    if failed_sheets:
        msg += "\n\nFailed/Skipped:\n" + "\n".join(failed_sheets[:10])
        if len(failed_sheets) > 10:
            msg += "\n... and {} more".format(len(failed_sheets) - 10)
    UI.TaskDialog.Show("Sheet Scale Updater", msg)

    _print_md("# Sheet Scale Updater - Detailed Report\n")
    _print_md("**Developed by: Jason Tian**\n\n")
    _print_md("**Updated:** {} sheets\n".format(updated_count))
    _print_md("**Failed/Skipped:** {} sheets\n".format(len(failed_sheets)))
    _print_md("**Target Parameter:** `{}` ({})\n".format(target_param_name, target_param_scope or "instance"))
    _print_md("\n## Per-Sheet Details:\n")

    for debug in debug_info:
        _print_md("\n### Sheet: `{}`".format(debug.get("sheet", "Unknown")))
        if debug.get("status") == "SUCCESS":
            _print_md("**Status:** [SUCCESS]")
        else:
            _print_md("**Status:** [FAILED]")
        if debug.get("titleblock_found"):
            _print_md("- **Titleblock:** Found")
        if "error" in debug and "titleblock" in debug["error"].lower():
            _print_md("- **Titleblock:** NOT FOUND [WARNING]")
        if "viewport_count" in debug:
            _print_md("- **Viewports:** {} found".format(debug["viewport_count"]))
        if "all_viewport_scales" in debug:
            _print_md("- **All Viewport Scales (read):** {}".format(debug["all_viewport_scales"]))
        if "scales_read" in debug:
            _print_md("- **Valid Scales (after filter):** {}".format(debug["scales_read"]))
        if "scale_selection" in debug:
            _print_md("- **Scale Selection Method:** {}".format(debug["scale_selection"]))
        if debug.get("multiple_scales"):
            _print_md("- **Multiple Scales Detected:** Yes")
            if "override_param_exists" in debug:
                _print_md("  - **Override Parameter Exists:** {}".format(debug["override_param_exists"]))
            if "override_value" in debug:
                _print_md("  - **Override Value:** {}".format(debug["override_value"]))
        if "scale_to_write" in debug:
            _print_md("- **Scale to Write:** {}".format(debug["scale_to_write"]))
        if "sheet_scale_param_exists" in debug:
            _print_md("- **Sheet Scale Parameter Exists:** {}".format(debug["sheet_scale_param_exists"]))
        if "param_storage_type" in debug:
            _print_md("- **Parameter Storage Type:** {}".format(debug["param_storage_type"]))
        if "written_as" in debug:
            _print_md("- **Written As:** {}".format(debug["written_as"]))
        if "error" in debug:
            _print_md("- **Error:** {}".format(debug["error"]))

main()
