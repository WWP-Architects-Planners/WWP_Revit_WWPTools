#! python3
import collections
from collections import defaultdict
import os

try:
    from collections.abc import Callable as _Callable
except Exception:
    _Callable = None

if _Callable is not None and not hasattr(collections, "Callable"):
    collections.Callable = _Callable
if not hasattr(collections, "callable"):
    collections.callable = callable

from System.Collections.Generic import List

import clr

import sys

from pyrevit import DB, revit, script

clr.AddReference("RevitAPIUI")
from Autodesk.Revit import UI


def _load_uiutils():
    script_dir = os.path.dirname(__file__)
    lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))
    if lib_path not in sys.path:
        sys.path.append(lib_path)
    import WWP_uiUtils as ui
    return ui


def _pick_import_instance(uidoc):
    sel_ids = list(uidoc.Selection.GetElementIds())
    if len(sel_ids) == 1:
        elem = uidoc.Document.GetElement(sel_ids[0])
        if isinstance(elem, DB.ImportInstance):
            return elem

    class _ImportInstanceFilter(DB.Selection.ISelectionFilter):
        def AllowElement(self, element):
            return isinstance(element, DB.ImportInstance)

        def AllowReference(self, reference, position):
            return False

    try:
        ref = uidoc.Selection.PickObject(
            DB.Selection.ObjectType.Element,
            _ImportInstanceFilter(),
            "Pick a CAD import/link",
        )
    except Exception:
        return None

    return uidoc.Document.GetElement(ref.ElementId)


def _choose_import_instance(doc, view, ui):
    uidoc = revit.uidoc
    picked = _pick_import_instance(uidoc)
    if picked:
        return picked

    in_view = list(
        DB.FilteredElementCollector(doc, view.Id)
        .OfClass(DB.ImportInstance)
        .WhereElementIsNotElementType()
    )
    in_doc = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.ImportInstance)
        .WhereElementIsNotElementType()
    )
    candidates = in_view or in_doc
    if not candidates:
        return None

    options = {e.Name: e for e in candidates}
    options_list = sorted(options.keys())
    chosen = _select_single("Select CAD Import", options_list, ui)
    if not chosen:
        return None
    return options[chosen]


def _get_layer_names(import_instance):
    cat = import_instance.Category
    if cat and cat.SubCategories:
        return sorted([sc.Name for sc in cat.SubCategories])
    return []


def _walk_geometry(geom_elem, transform):
    for obj in geom_elem:
        if isinstance(obj, DB.GeometryInstance):
            next_t = transform.Multiply(obj.Transform) if transform else obj.Transform
            if obj.SymbolGeometry is None:
                continue
            for sub in _walk_geometry(obj.SymbolGeometry, next_t):
                yield sub
        else:
            yield obj, transform


def _get_layer_name(doc, geom_obj):
    if geom_obj is None:
        return None
    try:
        gs_id = geom_obj.GraphicsStyleId
    except Exception:
        return None
    if gs_id == DB.ElementId.InvalidElementId:
        return None
    gs = doc.GetElement(gs_id)
    if isinstance(gs, DB.GraphicsStyle) and gs.GraphicsStyleCategory:
        return gs.GraphicsStyleCategory.Name
    return None


def _polyline_to_lines(polyline, transform, tol):
    pts = list(polyline.GetCoordinates())
    if transform:
        pts = [transform.OfPoint(p) for p in pts]
    lines = []
    for i in range(len(pts) - 1):
        try:
            if pts[i].DistanceTo(pts[i + 1]) < tol:
                continue
            lines.append(DB.Line.CreateBound(pts[i], pts[i + 1]))
        except Exception:
            continue
    return lines


def _collect_curves_by_layer(doc, import_instance, selected_layers):
    opts = DB.Options()
    opts.IncludeNonVisibleObjects = True
    opts.DetailLevel = DB.ViewDetailLevel.Fine

    geom = import_instance.get_Geometry(opts)
    if not geom:
        return {}
    curves_by_layer = defaultdict(list)

    tol = doc.Application.ShortCurveTolerance
    for obj, transform in _walk_geometry(geom, None):
        if obj is None:
            continue
        layer_name = _get_layer_name(doc, obj)
        if not layer_name or layer_name not in selected_layers:
            continue

        if isinstance(obj, DB.Curve):
            curve = obj.CreateTransformed(transform) if transform else obj
            if curve.Length >= tol:
                curves_by_layer[layer_name].append(curve)
        elif isinstance(obj, DB.PolyLine):
            for seg in _polyline_to_lines(obj, transform, tol):
                if seg.Length >= tol:
                    curves_by_layer[layer_name].append(seg)

    return curves_by_layer


def _unique_group_type_name(doc, base_name):
    existing = {
        gt.Name
        for gt in DB.FilteredElementCollector(doc)
        .OfClass(DB.GroupType)
        .ToElements()
    }
    if base_name not in existing:
        return base_name

    idx = 1
    while True:
        candidate = "{} ({})".format(base_name, idx)
        if candidate not in existing:
            return candidate
        idx += 1


def main():
    doc = revit.doc
    view = doc.ActiveView
    ui = _load_uiutils()
    if view.ViewType in (DB.ViewType.ThreeD, DB.ViewType.DrawingSheet):
        _show_alert("Detail Lines", "Open a plan/section/detail view before running.", ui)
        return

    import_instance = _choose_import_instance(doc, view, ui)
    if not import_instance:
        _show_alert("Detail Lines", "No CAD import/link selected.", ui)
        return

    layer_names = _get_layer_names(import_instance)
    if not layer_names:
        _show_alert("Detail Lines", "No layers found on the selected CAD import.", ui)
        return

    selected_layers = _select_multi("Select CAD Layers", layer_names, ui)
    if not selected_layers:
        return

    line_styles = _get_line_styles(doc)
    if not line_styles:
        _show_alert("Detail Lines", "No line styles found in this document.", ui)
        return

    style_names = _order_line_style_names(line_styles.keys())
    layer_style_map = _map_layers_to_styles(selected_layers, style_names, ui)
    if not layer_style_map:
        return

    curves_by_layer = _collect_curves_by_layer(doc, import_instance, set(selected_layers))
    if not curves_by_layer:
        _show_alert("Detail Lines", "No curves found for the selected layers.", ui)
        return

    created_counts = {}
    skipped = 0
    skipped_short = 0
    ungrouped = 0
    tol = doc.Application.ShortCurveTolerance

    with revit.Transaction("Detail Lines from CAD"):
        for layer_name, curves in curves_by_layer.items():
            ids = List[DB.ElementId]()
            for curve in curves:
                try:
                    if curve.Length < tol:
                        skipped_short += 1
                        continue
                except Exception:
                    pass
                try:
                    detail = doc.Create.NewDetailCurve(view, curve)
                except Exception as exc:
                    if "ShortCurveTolerance" in str(exc) or "endpoints" in str(exc):
                        skipped_short += 1
                        continue
                    skipped += 1
                    continue
                style_name = layer_style_map.get(layer_name)
                if style_name:
                    style = line_styles.get(style_name)
                    if style:
                        try:
                            detail.LineStyle = style
                        except Exception:
                            pass
                ids.Add(detail.Id)

            if ids.Count > 1:
                try:
                    group = doc.Create.NewGroup(ids)
                    group.GroupType.Name = _unique_group_type_name(doc, layer_name)
                except Exception:
                    ungrouped += ids.Count
                created_counts[layer_name] = ids.Count
            elif ids.Count == 1:
                ungrouped += 1
                created_counts[layer_name] = 1

    _report_results(created_counts, skipped, ungrouped, skipped_short)


def _show_alert(title, message, ui):
    ui.uiUtils_alert(message, title=title)


def _get_line_styles(doc):
    styles = {}
    try:
        cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
    except Exception:
        return styles

    try:
        gs = cat.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
        if gs:
            styles[cat.Name] = gs
    except Exception:
        pass

    try:
        for subcat in cat.SubCategories:
            try:
                gs = subcat.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
                if gs:
                    styles[subcat.Name] = gs
            except Exception:
                continue
    except Exception:
        pass

    return styles


def _order_line_style_names(style_names):
    names = sorted(style_names)
    if "Lines" in names:
        names.remove("Lines")
        names.insert(0, "Lines")
    return names


def _map_layers_to_styles(layers, style_names, ui):
    if not layers or not style_names:
        return None
    use_single = ui.uiUtils_confirm(
        "Use a single line style for all selected layers?",
        title="Map CAD Layers to Line Styles",
    )
    if use_single:
        selected = _select_single("Select Line Style", style_names, ui)
        if not selected:
            return None
        return {layer: selected for layer in layers}

    mapping = {}
    for layer in layers:
        style = _select_single(
            "Map Layer: {}".format(layer),
            style_names,
            ui,
        )
        if not style:
            return None
        mapping[layer] = style
    return mapping


def _report_results(created_counts, skipped, ungrouped, skipped_short):
    lines = ["Detail Lines from CAD"]
    for lname in sorted(created_counts.keys()):
        lines.append("- {}: {} detail lines".format(lname, created_counts[lname]))
    if skipped:
        lines.append("- Skipped: {} curves (not in view plane or invalid)".format(skipped))
    if ungrouped:
        lines.append("- Ungrouped: {} detail lines (single or group error)".format(ungrouped))
    if skipped_short:
        lines.append("- Skipped (short): {} curves below tolerance".format(skipped_short))

    try:
        out = script.get_output()
        if hasattr(sys.stdout, "flush"):
            out.print_md("**Detail Lines from CAD**")
            for line in lines[1:]:
                out.print_md(line)
            return
    except Exception:
        pass

    for line in lines:
        print(line)


def _select_single(title, options, ui):
    if not options:
        return None
    indices = ui.uiUtils_select_indices(
        options,
        title=title,
        prompt="Select an option:",
        multiselect=False,
        width=520,
        height=540,
    )
    if not indices:
        return None
    index = indices[0]
    if 0 <= index < len(options):
        return options[index]
    return None


def _select_multi(title, options, ui):
    if not options:
        return []
    indices = ui.uiUtils_select_indices(
        options,
        title=title,
        prompt="Select items:",
        multiselect=True,
        width=520,
        height=540,
    )
    return [options[i] for i in indices if 0 <= i < len(options)]


if __name__ == "__main__":
    main()
