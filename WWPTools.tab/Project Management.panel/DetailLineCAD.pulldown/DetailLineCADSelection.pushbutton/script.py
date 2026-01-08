#! python3
from collections import defaultdict

from System.Collections.Generic import List

from pyrevit import DB, revit, forms, script


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


def _choose_import_instance(doc, view):
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
    chosen = forms.SelectFromList.show(
        sorted(options.keys()),
        multiselect=False,
        title="Select CAD Import",
        button_name="Select",
    )
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
            for sub in _walk_geometry(obj.SymbolGeometry, next_t):
                yield sub
        else:
            yield obj, transform


def _get_layer_name(doc, geom_obj):
    gs_id = geom_obj.GraphicsStyleId
    if gs_id == DB.ElementId.InvalidElementId:
        return None
    gs = doc.GetElement(gs_id)
    if isinstance(gs, DB.GraphicsStyle):
        return gs.GraphicsStyleCategory.Name
    return None


def _polyline_to_lines(polyline, transform):
    pts = list(polyline.GetCoordinates())
    if transform:
        pts = [transform.OfPoint(p) for p in pts]
    lines = []
    for i in range(len(pts) - 1):
        lines.append(DB.Line.CreateBound(pts[i], pts[i + 1]))
    return lines


def _collect_curves_by_layer(doc, import_instance, selected_layers):
    opts = DB.Options()
    opts.IncludeNonVisibleObjects = True
    opts.DetailLevel = DB.ViewDetailLevel.Fine

    geom = import_instance.get_Geometry(opts)
    curves_by_layer = defaultdict(list)

    for obj, transform in _walk_geometry(geom, None):
        layer_name = _get_layer_name(doc, obj)
        if not layer_name or layer_name not in selected_layers:
            continue

        if isinstance(obj, DB.Curve):
            curve = obj.CreateTransformed(transform) if transform else obj
            curves_by_layer[layer_name].append(curve)
        elif isinstance(obj, DB.PolyLine):
            curves_by_layer[layer_name].extend(
                _polyline_to_lines(obj, transform)
            )

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
    if view.ViewType in (DB.ViewType.ThreeD, DB.ViewType.DrawingSheet):
        forms.alert("Open a plan/section/detail view before running.", title="Detail Lines")
        return

    import_instance = _choose_import_instance(doc, view)
    if not import_instance:
        forms.alert("No CAD import/link selected.", title="Detail Lines")
        return

    layer_names = _get_layer_names(import_instance)
    if not layer_names:
        forms.alert("No layers found on the selected CAD import.", title="Detail Lines")
        return

    selected_layers = forms.SelectFromList.show(
        layer_names,
        multiselect=True,
        title="Select CAD Layers",
        button_name="Convert",
    )
    if not selected_layers:
        return

    curves_by_layer = _collect_curves_by_layer(doc, import_instance, set(selected_layers))
    if not curves_by_layer:
        forms.alert("No curves found for the selected layers.", title="Detail Lines")
        return

    created_counts = {}
    skipped = 0

    with revit.Transaction("Detail Lines from CAD"):
        for layer_name, curves in curves_by_layer.items():
            ids = List[DB.ElementId]()
            for curve in curves:
                try:
                    detail = doc.Create.NewDetailCurve(view, curve)
                except Exception:
                    skipped += 1
                    continue
                ids.Add(detail.Id)

            if ids.Count:
                group = doc.Create.NewGroup(ids)
                group.GroupType.Name = _unique_group_type_name(doc, layer_name)
                created_counts[layer_name] = ids.Count

    out = script.get_output()
    out.print_md("**Detail Lines from CAD**")
    for lname in sorted(created_counts.keys()):
        out.print_md("- `{}`: {} detail lines".format(lname, created_counts[lname]))
    if skipped:
        out.print_md("- Skipped: {} curves (not in view plane or invalid)".format(skipped))


if __name__ == "__main__":
    main()
