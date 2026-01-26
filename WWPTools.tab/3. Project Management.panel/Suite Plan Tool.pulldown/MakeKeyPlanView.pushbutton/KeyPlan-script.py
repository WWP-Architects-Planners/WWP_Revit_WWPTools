#!python3
import clr

from pyrevit import revit, DB, script
import WWP_uiUtils as uiutils
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import Selection as UISelection
from System.Collections.Generic import List


doc = revit.doc
uidoc = revit.uidoc

CONFIG_LAST_AREA_ID = "last_area_id"
CONFIG_LAST_KEYPLAN_TEMPLATE_ID = "last_keyplan_template_id"
CONFIG_LAST_FILL_TYPE_ID = "last_fill_type_id"


def pick_elements(bic, prompt):
    bic_id = int(bic)
    view = doc.ActiveView
    isolate_active = False
    if bic_id == int(DB.BuiltInCategory.OST_Areas):
        try:
            view.IsolateCategoriesTemporary(List[DB.ElementId]([DB.ElementId(bic_id)]))
            isolate_active = True
        except Exception:
            isolate_active = False
    try:
        refs = uidoc.Selection.PickObjects(UISelection.ObjectType.Element, prompt)
        elems = []
        for ref in refs:
            elem = doc.GetElement(ref)
            if elem and elem.Category and elem.Category.Id.IntegerValue == bic_id:
                elems.append(elem)
        return elems
    finally:
        if isolate_active:
            try:
                view.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate)
            except Exception:
                pass


def element_id_value(elem_id):
    if elem_id is None:
        return None
    if hasattr(elem_id, "IntegerValue"):
        return elem_id.IntegerValue
    if hasattr(elem_id, "Value"):
        return elem_id.Value
    try:
        return int(elem_id)
    except Exception:
        return None


def element_is_category(elem, bic):
    if not elem or not elem.Category:
        return False
    return elem.Category.Id.IntegerValue == int(bic)


def unique_view_name(base_name):
    existing = set(v.Name for v in DB.FilteredElementCollector(doc).OfClass(DB.View))
    if base_name not in existing:
        return base_name
    index = 1
    while True:
        candidate = "{} ({})".format(base_name, index)
        if candidate not in existing:
            return candidate
        index += 1


def _safe_elem_label(elem):
    try:
        return elem.Name or elem.Id.IntegerValue
    except Exception:
        return "(unavailable)"


def curve_loop_area_xy(curves):
    pts = []
    for curve in curves:
        for pt in curve.Tessellate():
            if not pts or pt.DistanceTo(pts[-1]) > 1e-6:
                pts.append(pt)
    if len(pts) < 3:
        return 0.0
    if pts[0].DistanceTo(pts[-1]) > 1e-6:
        pts.append(pts[0])
    area = 0.0
    for i in range(len(pts) - 1):
        area += pts[i].X * pts[i + 1].Y - pts[i + 1].X * pts[i].Y
    return 0.5 * area


def get_outer_boundary_loop(area):
    opts = DB.SpatialElementBoundaryOptions()
    loops = area.GetBoundarySegments(opts)
    if not loops:
        return None
    best_curves = None
    best_area = -1.0
    for segs in loops:
        curves = [seg.GetCurve() for seg in segs]
        area_val = abs(curve_loop_area_xy(curves))
        if area_val > best_area:
            best_area = area_val
            best_curves = curves
    if not best_curves:
        return None
    loop = DB.CurveLoop()
    for curve in best_curves:
        try:
            loop.Append(curve)
        except Exception:
            return None
    return loop


def _rect_loop_from_bbox(bbox):
    if not bbox:
        return None
    min_pt = bbox.Min
    max_pt = bbox.Max
    if not min_pt or not max_pt:
        return None
    p1 = DB.XYZ(min_pt.X, min_pt.Y, min_pt.Z)
    p2 = DB.XYZ(max_pt.X, min_pt.Y, min_pt.Z)
    p3 = DB.XYZ(max_pt.X, max_pt.Y, min_pt.Z)
    p4 = DB.XYZ(min_pt.X, max_pt.Y, min_pt.Z)
    loop = DB.CurveLoop()
    loop.Append(DB.Line.CreateBound(p1, p2))
    loop.Append(DB.Line.CreateBound(p2, p3))
    loop.Append(DB.Line.CreateBound(p3, p4))
    loop.Append(DB.Line.CreateBound(p4, p1))
    return loop


def main():
    base_view = doc.ActiveView
    if not isinstance(base_view, DB.ViewPlan):
        uiutils.uiUtils_alert("Active view must be a plan view.", title="Make Keyplans")
        return

    templates = [
        v for v in DB.FilteredElementCollector(doc).OfClass(DB.View) if v.IsTemplate
    ]
    templates_sorted = sorted(templates, key=lambda v: v.Name)

    fill_types = (
        DB.FilteredElementCollector(doc)
        .OfClass(DB.FilledRegionType)
        .WhereElementIsElementType()
        .ToElements()
    )
    fill_types_sorted = sorted(fill_types, key=lambda f: f.Name)

    if not templates_sorted:
        uiutils.uiUtils_alert("No view templates found for this view type.", title="Make Keyplans")
        return
    if not fill_types_sorted:
        uiutils.uiUtils_alert("No filled region types found.", title="Make Keyplans")
        return

    config = script.get_config()
    state = {
        "areas": [],
        "area_label": None,
        "keyplan_template_index": 0,
        "fill_type_index": 0,
    }

    last_area_id = getattr(config, CONFIG_LAST_AREA_ID, None)
    if last_area_id:
        area_elem = doc.GetElement(DB.ElementId(int(last_area_id)))
        if element_is_category(area_elem, DB.BuiltInCategory.OST_Areas):
            state["areas"] = [area_elem]
            state["area_label"] = area_elem.Name or area_elem.Id.IntegerValue

    last_keyplan_template_id = getattr(config, CONFIG_LAST_KEYPLAN_TEMPLATE_ID, None)
    if last_keyplan_template_id:
        for idx, template in enumerate(templates_sorted):
            if template.Id.IntegerValue == int(last_keyplan_template_id):
                state["keyplan_template_index"] = idx
                break

    last_fill_type_id = getattr(config, CONFIG_LAST_FILL_TYPE_ID, None)
    if last_fill_type_id:
        for idx, fill_type in enumerate(fill_types_sorted):
            if fill_type.Id.IntegerValue == int(last_fill_type_id):
                state["fill_type_index"] = idx
                break

    if state["areas"]:
        use_last = uiutils.uiUtils_confirm("Use previously selected area(s)?", title="Make Keyplans")
        if not use_last:
            state["areas"] = []

    if not state["areas"]:
        try:
            areas = pick_elements(DB.BuiltInCategory.OST_Areas, "Select suite areas")
        except OperationCanceledException:
            return
        state["areas"] = areas

    areas = state.get("areas") or []
    if not areas:
        uiutils.uiUtils_alert("Please pick at least one area.", title="Make Keyplans")
        return

    if len(areas) > 1:
        state["area_label"] = "{} areas selected".format(len(areas))
    else:
        state["area_label"] = areas[0].Name or areas[0].Id.IntegerValue

    result = uiutils.uiUtils_keyplan_options(
        [t.Name for t in templates_sorted],
        [f.Name for f in fill_types_sorted],
        title="Make Keyplans",
        area_label=state.get("area_label") or "(not selected)",
        template_index=state.get("keyplan_template_index", 0),
        fill_type_index=state.get("fill_type_index", 0),
        width=640,
        height=360,
    )
    if result is None:
        return

    template_index = result.get("template_index", 0)
    fill_type_index = result.get("fill_type_index", 0)

    keyplan_template = (
        templates_sorted[template_index]
        if 0 <= template_index < len(templates_sorted)
        else templates_sorted[0]
    )
    keyplan_fill_type = (
        fill_types_sorted[fill_type_index]
        if 0 <= fill_type_index < len(fill_types_sorted)
        else fill_types_sorted[0]
    )

    area_for_config = areas[0] if areas else None
    config.last_area_id = element_id_value(area_for_config.Id) if area_for_config else None
    config.last_keyplan_template_id = (
        element_id_value(keyplan_template.Id) if keyplan_template else None
    )
    config.last_fill_type_id = (
        element_id_value(keyplan_fill_type.Id) if keyplan_fill_type else None
    )
    script.save_config()

    results = []
    with revit.Transaction("Create Keyplan Views"):
        for area in areas:
            loop = get_outer_boundary_loop(area)
            fallback_loop = None
            if not loop:
                fallback_loop = _rect_loop_from_bbox(area.get_BoundingBox(base_view))
                loop = fallback_loop
            if not loop:
                results.append(
                    {
                        "area": area,
                        "keyplan_view": None,
                        "warnings": ["Could not create a filled region boundary."],
                        "failed": True,
                    }
                )
                continue

            base_label = area.Number or area.Name or str(area.Id.IntegerValue)
            keyplan_view_id = base_view.Duplicate(DB.ViewDuplicateOption.Duplicate)
            keyplan_view = doc.GetElement(keyplan_view_id)
            keyplan_view.Name = unique_view_name("Keyplan - {}".format(base_label))
            keyplan_view.ViewTemplateId = keyplan_template.Id
            keyplan_view.CropBoxActive = True
            keyplan_view.CropBoxVisible = True

            create_errors = []
            fill_loops = List[DB.CurveLoop]()
            try:
                fill_loops.Add(loop)
                DB.FilledRegion.Create(doc, keyplan_fill_type.Id, keyplan_view.Id, fill_loops)
            except Exception:
                if fallback_loop is None:
                    fallback_loop = _rect_loop_from_bbox(area.get_BoundingBox(base_view))
                if fallback_loop:
                    try:
                        fill_loops = List[DB.CurveLoop]()
                        fill_loops.Add(fallback_loop)
                        DB.FilledRegion.Create(
                            doc, keyplan_fill_type.Id, keyplan_view.Id, fill_loops
                        )
                        create_errors.append(
                            "Filled region loop was discontinuous; used area bounding box."
                        )
                    except Exception:
                        create_errors.append("Filled region could not be created.")
                else:
                    create_errors.append("Filled region could not be created.")

            results.append(
                {
                    "area": area,
                    "keyplan_view": keyplan_view,
                    "warnings": create_errors,
                    "failed": False,
                }
            )

    last_keyplan = None
    for result in results:
        if result.get("keyplan_view"):
            last_keyplan = result.get("keyplan_view")

    if last_keyplan:
        try:
            uidoc.RequestViewChange(last_keyplan)
        except Exception:
            uidoc.ActiveView = last_keyplan

    created = [r for r in results if r.get("keyplan_view")]
    failed = [r for r in results if r.get("failed")]
    warnings = []
    for result in results:
        for warning in result.get("warnings") or []:
            warnings.append(warning)

    lines = ["Created {} keyplan view(s):".format(len(created))]
    for result in created:
        view = result.get("keyplan_view")
        area = result.get("area")
        lines.append("- {} ({})".format(view.Name, _safe_elem_label(area)))
    if failed:
        lines.append("")
        lines.append("Failed {} area(s):".format(len(failed)))
        for result in failed:
            area = result.get("area")
            lines.append("- {}".format(_safe_elem_label(area)))
    message = "\n".join(lines)
    if warnings:
        message += "\n\nWarnings:\n" + "\n".join(sorted(set(warnings)))
    uiutils.uiUtils_alert(message, title="Keyplans Created")


if __name__ == "__main__":
    main()
