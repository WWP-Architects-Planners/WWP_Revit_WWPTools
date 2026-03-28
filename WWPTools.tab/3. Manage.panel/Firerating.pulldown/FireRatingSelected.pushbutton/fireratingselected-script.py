#!python3
# -*- coding: utf-8 -*-
"""Create Fire Rating Lines for Selected Walls.

Operates on the currently selected walls in the active view,
applying line styles based on the FRR Walls parameter value.
"""
import os
import sys
import traceback

script_dir = os.path.dirname(__file__)
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "..", "lib"))
if lib_path not in sys.path:
    sys.path.append(lib_path)

import WWP_uiUtils as ui
from pyrevit import DB

uidoc = getattr(__revit__, "ActiveUIDocument", None)
doc = uidoc.Document if uidoc else None
WINDOW_TITLE = "Create Fire Rating Lines — Selected Walls"
FRR_PARAM_NAME = "FRR Walls"


def get_all_line_styles():
    lines_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
    return {sub.Name: sub.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
            for sub in lines_cat.SubCategories}


def find_style_for_frr(frr_value, all_styles):
    if not frr_value:
        return None
    frr_upper = frr_value.strip().upper()
    for name, style in all_styles.items():
        if "FRR" in name.upper() and frr_upper in name.upper():
            return style
    frr_dash = frr_upper.replace("/", "-")
    for name, style in all_styles.items():
        if "FRR" in name.upper() and frr_dash in name.upper():
            return style
    return None


def main():
    if doc is None:
        ui.uiUtils_alert("Open a project first.", title=WINDOW_TITLE)
        return

    active_view = doc.ActiveView
    sel_ids = list(uidoc.Selection.GetElementIds())
    walls = [
        doc.GetElement(eid)
        for eid in sel_ids
        if isinstance(doc.GetElement(eid), DB.Wall)
    ]

    if not walls:
        ui.uiUtils_alert(
            "No walls selected. Please select one or more walls first.",
            title=WINDOW_TITLE,
        )
        return

    all_styles = get_all_line_styles()
    created = 0
    skipped_no_param = 0
    skipped_zero = 0

    t = DB.Transaction(doc, "Create Fire Rating Lines (Selected)")
    t.Start()
    try:
        for wall in walls:
            loc = wall.Location
            if not isinstance(loc, DB.LocationCurve):
                continue
            curve = loc.Curve

            param = wall.LookupParameter(FRR_PARAM_NAME)
            if param is None:
                skipped_no_param += 1
                continue

            frr_val = ""
            if param.StorageType == DB.StorageType.String:
                frr_val = param.AsString() or ""

            if not frr_val or frr_val.upper() == "0HR":
                skipped_zero += 1
                continue

            try:
                detail_line = doc.Create.NewDetailCurve(active_view, curve)
                style = find_style_for_frr(frr_val, all_styles)
                if style is not None:
                    detail_line.LineStyle = style
                created += 1
            except Exception:
                pass
        t.Commit()
    except Exception:
        t.RollBack()
        raise

    lines = ["Created {} fire rating line(s) from {} selected wall(s).".format(
        created, len(walls))]
    if skipped_zero:
        lines.append("{} wall(s) skipped — rated 0HR or no rating.".format(skipped_zero))
    if skipped_no_param:
        lines.append("{} wall(s) skipped — missing '{}' parameter.".format(
            skipped_no_param, FRR_PARAM_NAME))
    ui.uiUtils_alert("\n".join(lines), title=WINDOW_TITLE)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        ui.uiUtils_alert(traceback.format_exc(), title=WINDOW_TITLE)
