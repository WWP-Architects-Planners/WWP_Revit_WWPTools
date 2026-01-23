#!python3
import clr

from pyrevit import revit, DB, script
import WWP_uiUtils as uiutils
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import Selection as UISelection

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Drawing import Point, Size
from System.Collections.Generic import List
from System.Windows.Forms import (
    Button,
    ComboBox,
    ComboBoxStyle,
    DialogResult,
    Form,
    FormStartPosition,
    Label,
)


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


class KeyplanForm(Form):
    def __init__(self, base_view, templates, fill_types, state):
        super(KeyplanForm, self).__init__()
        self.Text = "Make Keyplans"
        self.StartPosition = FormStartPosition.CenterScreen
        self.Size = Size(640, 320)
        self.MinimizeBox = False
        self.MaximizeBox = False

        self._base_view = base_view
        self._templates = templates
        self._fill_types = fill_types
        self._state = state

        self.selected_keyplan_template = None
        self.selected_fill_type = None
        self.request_action = None

        current_y = 16

        info = Label()
        info.Text = "Active view: {}".format(base_view.Name)
        info.Location = Point(12, current_y)
        info.Size = Size(600, 20)
        self.Controls.Add(info)

        current_y += 32
        self._area_label = Label()
        area_label = self._state.get("area_label") or "(not selected)"
        self._area_label.Text = "Areas: {}".format(area_label)
        self._area_label.Location = Point(12, current_y)
        self._area_label.Size = Size(420, 20)
        self.Controls.Add(self._area_label)

        self._area_btn = Button()
        self._area_btn.Text = "Pick Areas"
        self._area_btn.Location = Point(460, current_y - 2)
        self._area_btn.Size = Size(140, 26)
        self._area_btn.Click += self._on_pick_area
        self.Controls.Add(self._area_btn)

        current_y += 40
        template_label = Label()
        template_label.Text = "Keyplan view template:"
        template_label.Location = Point(12, current_y)
        template_label.Size = Size(220, 20)
        self.Controls.Add(template_label)

        self._keyplan_template_combo = ComboBox()
        self._keyplan_template_combo.Location = Point(240, current_y - 2)
        self._keyplan_template_combo.Size = Size(360, 26)
        self._keyplan_template_combo.DropDownStyle = ComboBoxStyle.DropDownList
        self.Controls.Add(self._keyplan_template_combo)

        current_y += 36
        fill_label = Label()
        fill_label.Text = "Keyplan filled region type:"
        fill_label.Location = Point(12, current_y)
        fill_label.Size = Size(220, 20)
        self.Controls.Add(fill_label)

        self._fill_combo = ComboBox()
        self._fill_combo.Location = Point(240, current_y - 2)
        self._fill_combo.Size = Size(360, 26)
        self._fill_combo.DropDownStyle = ComboBoxStyle.DropDownList
        self.Controls.Add(self._fill_combo)

        current_y += 40
        self._create_btn = Button()
        self._create_btn.Text = "Create"
        self._create_btn.Location = Point(420, current_y)
        self._create_btn.Size = Size(90, 30)
        self._create_btn.Click += self._on_create
        self.Controls.Add(self._create_btn)

        self._cancel_btn = Button()
        self._cancel_btn.Text = "Cancel"
        self._cancel_btn.Location = Point(520, current_y)
        self._cancel_btn.Size = Size(90, 30)
        self._cancel_btn.Click += self._on_cancel
        self.Controls.Add(self._cancel_btn)

        self._load_template_choices()
        self._load_fill_choices()

    def _load_template_choices(self):
        self._keyplan_template_combo.Items.Clear()
        for template in self._templates:
            self._keyplan_template_combo.Items.Add(template.Name)
        if self._templates:
            index = self._state.get("keyplan_template_index", 0)
            if isinstance(index, int) and 0 <= index < len(self._templates):
                self._keyplan_template_combo.SelectedIndex = index
            else:
                self._keyplan_template_combo.SelectedIndex = 0

    def _load_fill_choices(self):
        self._fill_combo.Items.Clear()
        for fill_type in self._fill_types:
            self._fill_combo.Items.Add(fill_type.Name)
        if self._fill_types:
            index = self._state.get("fill_type_index", 0)
            if isinstance(index, int) and 0 <= index < len(self._fill_types):
                self._fill_combo.SelectedIndex = index
            else:
                self._fill_combo.SelectedIndex = 0

    def _on_pick_area(self, sender, args):
        if self._keyplan_template_combo.SelectedIndex >= 0:
            self._state["keyplan_template_index"] = self._keyplan_template_combo.SelectedIndex
        if self._fill_combo.SelectedIndex >= 0:
            self._state["fill_type_index"] = self._fill_combo.SelectedIndex
        self.request_action = "area"
        self.DialogResult = DialogResult.Retry
        self.Close()

    def _on_create(self, sender, args):
        if not self._state.get("areas"):
            uiutils.uiUtils_alert("Please pick at least one area.")
            return
        if self._keyplan_template_combo.SelectedIndex < 0:
            uiutils.uiUtils_alert("Please select a keyplan view template.")
            return
        if self._fill_combo.SelectedIndex < 0:
            uiutils.uiUtils_alert("Please select a keyplan filled region type.")
            return

        self.selected_keyplan_template = self._templates[self._keyplan_template_combo.SelectedIndex]
        self.selected_fill_type = self._fill_types[self._fill_combo.SelectedIndex]

        self.DialogResult = DialogResult.OK
        self.Close()

    def _on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


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

    while True:
        form = KeyplanForm(base_view, templates_sorted, fill_types_sorted, state)
        result = form.ShowDialog()
        if result == DialogResult.Retry:
            try:
                areas = pick_elements(DB.BuiltInCategory.OST_Areas, "Select suite areas")
            except OperationCanceledException:
                continue
            if not areas:
                state["areas"] = []
                state["area_label"] = "(not selected)"
            else:
                state["areas"] = areas
                if len(areas) > 1:
                    state["area_label"] = "{} areas selected".format(len(areas))
                else:
                    state["area_label"] = areas[0].Name or areas[0].Id.IntegerValue
            continue
        if result != DialogResult.OK:
            return

        areas = state.get("areas") or []
        keyplan_template = form.selected_keyplan_template
        keyplan_fill_type = form.selected_fill_type

        area_for_config = areas[0] if areas else None
        config.last_area_id = element_id_value(area_for_config.Id) if area_for_config else None
        config.last_keyplan_template_id = (
            element_id_value(keyplan_template.Id) if keyplan_template else None
        )
        config.last_fill_type_id = (
            element_id_value(keyplan_fill_type.Id) if keyplan_fill_type else None
        )
        script.save_config()
        break

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
