#! python3
import math
import os
import sys
import traceback
from datetime import datetime

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit import UI
from pyrevit import DB
from pyrevit import revit
from pyrevit.framework import EventHandler
from System.Collections.Generic import List
from System.IO import File
from System.Windows import MessageBox, MessageBoxButton
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Markup import XamlReader
from System.Windows import Visibility


TITLE = "Round Off-Axis Sketch Lines"
TARGET_WARNING_TEXT = "Line in Sketch is slightly off axis and may cause inaccuracies."
ANGLE_PRECISION = 2
ANGLE_EPSILON_DEGREES = 1e-6
MAX_DETAIL_LINES = 60
WINDOW_TITLE = TITLE
_WPFUI_THEME_READY = False


SCRIPT_DIR = os.path.dirname(__file__)
LIB_PATH = os.path.abspath(os.path.join(SCRIPT_DIR, "..", "..", "..", "lib"))
if LIB_PATH not in sys.path:
    sys.path.append(LIB_PATH)

import WWP_uiUtils as ui
from WWP_versioning import apply_window_title


def _element_id_value(elem_id):
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


def _normalize_text(value):
    return " ".join(str(value or "").split()).strip()


def ensure_wpfui_theme():
    global _WPFUI_THEME_READY
    if _WPFUI_THEME_READY:
        return

    try:
        revit_version = int(str(__revit__.Application.VersionNumber))
    except Exception:
        revit_version = None

    dll_name = "WWPTools.WpfUI.net8.0-windows.dll" if revit_version and revit_version >= 2025 else "WWPTools.WpfUI.net48.dll"
    dll_path = os.path.join(LIB_PATH, dll_name)
    if not os.path.isfile(dll_path):
        return

    try:
        if hasattr(clr, "AddReferenceToFileAndPath"):
            clr.AddReferenceToFileAndPath(dll_path)
        else:
            clr.AddReference(dll_path)
        _WPFUI_THEME_READY = True
    except Exception:
        pass


def _set_owner(window):
    try:
        helper = WindowInteropHelper(window)
        helper.Owner = revit.uidoc.Application.MainWindowHandle
    except Exception:
        pass


def _normalize_angle_degrees(angle_degrees):
    value = float(angle_degrees)
    while value < 0.0:
        value += 360.0
    while value >= 360.0:
        value -= 360.0
    return value


def _shortest_delta_degrees(current_degrees, target_degrees):
    delta = _normalize_angle_degrees(target_degrees) - _normalize_angle_degrees(current_degrees)
    while delta <= -180.0:
        delta += 360.0
    while delta > 180.0:
        delta -= 360.0
    return delta


def _category_name(element):
    try:
        if element.Category and element.Category.Name:
            return element.Category.Name
    except Exception:
        pass
    return type(element).__name__


def _element_name(element):
    try:
        name = element.Name
        if name:
            return str(name)
    except Exception:
        pass
    return "<unnamed>"


def _element_label(element):
    return "{} {} ({})".format(
        _category_name(element),
        _element_id_value(getattr(element, "Id", None)),
        _element_name(element),
    )


def _level_name(doc, element):
    if doc is None or element is None:
        return ""
    for attr_name in ("LevelId", "GenLevel", "ReferenceLevel"):
        try:
            value = getattr(element, attr_name, None)
            if value is None:
                continue
            if hasattr(value, "Name"):
                return str(value.Name or "").strip()
            level = doc.GetElement(value)
            if level is not None and getattr(level, "Name", None):
                return str(level.Name).strip()
        except Exception:
            pass
    return ""


def _owner_view_name(doc, element):
    if doc is None or element is None:
        return ""
    for attr_name in ("OwnerViewId", "ViewSpecific"):
        try:
            if attr_name == "ViewSpecific" and not bool(getattr(element, attr_name, False)):
                continue
            if attr_name == "OwnerViewId":
                owner_view_id = getattr(element, attr_name, None)
                if owner_view_id is None:
                    continue
                owner_view = doc.GetElement(owner_view_id)
                if owner_view is not None and getattr(owner_view, "Name", None):
                    return str(owner_view.Name).strip()
        except Exception:
            pass
    return ""


def _owner_view(doc, element):
    if doc is None or element is None:
        return None
    try:
        owner_view_id = getattr(element, "OwnerViewId", None)
        if owner_view_id is None:
            return None
        view = doc.GetElement(owner_view_id)
        if isinstance(view, DB.View):
            return view
    except Exception:
        pass
    return None


def _is_navigable_view(view):
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        return False
    if isinstance(view, DB.ViewSheet):
        return False
    try:
        if not view.CanBePrinted:
            return False
    except Exception:
        return False
    try:
        rejected = {
            "ProjectBrowser",
            "SystemBrowser",
            "Schedule",
            "DrawingSheet",
            "Internal",
            "Undefined",
            "Report",
            "CostReport",
            "LoadsReport",
            "ColumnSchedule",
            "PanelSchedule",
        }
        if str(view.ViewType) in rejected:
            return False
    except Exception:
        pass
    return True


def _same_level_view(doc, element):
    if doc is None or element is None:
        return None
    target_level = _level_name(doc, element)
    if not target_level:
        return None
    matches = []
    for view in DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements():
        if not _is_navigable_view(view):
            continue
        if _level_name(doc, view) != target_level:
            continue
        matches.append(view)
    if not matches:
        return None
    matches.sort(key=lambda item: "{}|{}".format(str(getattr(item, "ViewType", "")), str(getattr(item, "Name", ""))).lower())
    return matches[0]


def _context_label(doc, element):
    label = _element_label(element)
    level_name = _level_name(doc, element)
    view_name = _owner_view_name(doc, element)
    extras = []
    if level_name:
        extras.append("Level: {}".format(level_name))
    if view_name:
        extras.append("View: {}".format(view_name))
    if extras:
        return "{} | {}".format(label, " | ".join(extras))
    return label


def _group_label(group_element):
    if group_element is None:
        return ""
    return _element_label(group_element)


def _resolved_group_element(record):
    if record is None:
        return None
    explicit_group = record.get("group_element")
    if explicit_group is not None:
        return explicit_group
    return _group_element(
        revit.doc,
        _top_parent_element(
            revit.doc,
            primary_element=record.get("element"),
            context_element=record.get("context_element"),
            explicit_group=None,
        ),
    )


def _append_group_to_detail(detail_text, group_element):
    group_text = _group_label(group_element)
    if not group_text:
        return detail_text
    return "{}\nHost group: {}".format(detail_text, group_text)


def _warning_context_text(doc, failing_elements, target_element):
    context_labels = []
    target_id = _element_id_value(getattr(target_element, "Id", None))
    for element in list(failing_elements or []):
        if element is None:
            continue
        if _element_id_value(getattr(element, "Id", None)) == target_id:
            continue
        label = _context_label(doc, element)
        if label not in context_labels:
            context_labels.append(label)

    if not context_labels:
        target_level = _level_name(doc, target_element)
        target_view = _owner_view_name(doc, target_element)
        extras = []
        if target_level:
            extras.append("Level: {}".format(target_level))
        if target_view:
            extras.append("View: {}".format(target_view))
        if extras:
            return " | ".join(extras)
        return ""

    if len(context_labels) <= 3:
        return " | ".join(context_labels)
    return " | ".join(context_labels[:3]) + " | +{} more".format(len(context_labels) - 3)


def _warning_context_element(failing_elements, target_element):
    target_id = _element_id_value(getattr(target_element, "Id", None))
    for element in list(failing_elements or []):
        if element is None:
            continue
        if _element_id_value(getattr(element, "Id", None)) == target_id:
            continue
        return element
    return None


def _is_property_line(element):
    try:
        category = element.Category
        if category is None:
            return False
        return category.Id.IntegerValue == int(DB.BuiltInCategory.OST_PropertyLines)
    except Exception:
        return False


def _invalid_element_id_value():
    invalid = getattr(DB.ElementId, "InvalidElementId", None)
    return _element_id_value(invalid)


def _is_group_member(element):
    try:
        group_id = getattr(element, "GroupId", None)
        if group_id is None:
            return False
        group_value = _element_id_value(group_id)
        invalid_value = _invalid_element_id_value()
        return group_value is not None and group_value != invalid_value
    except Exception:
        return False


def _group_element(doc, element):
    if doc is None or element is None:
        return None
    try:
        group_id = getattr(element, "GroupId", None)
        if group_id is None:
            return None
        group_value = _element_id_value(group_id)
        invalid_value = _invalid_element_id_value()
        if group_value is None or group_value == invalid_value:
            return None
        return doc.GetElement(group_id)
    except Exception:
        return None


def _top_parent_element(doc, primary_element=None, context_element=None, explicit_group=None):
    if explicit_group is not None:
        return explicit_group

    for candidate in (context_element, primary_element):
        if candidate is None:
            continue
        group = _group_element(doc, candidate)
        if group is not None:
            return group

    if context_element is not None:
        return context_element
    return primary_element


def _warning_description(warning):
    for attr_name in ("GetDescriptionText", "DescriptionText"):
        try:
            value = getattr(warning, attr_name)
            if callable(value):
                value = value()
            text = _normalize_text(value)
            if text:
                return text
        except Exception:
            pass
    return ""


def _collect_target_warnings(doc):
    warnings = []
    try:
        doc_warnings = list(doc.GetWarnings() or [])
    except Exception:
        doc_warnings = []

    target_text = _normalize_text(TARGET_WARNING_TEXT)
    for warning in doc_warnings:
        if _warning_description(warning) == target_text:
            warnings.append(warning)
    return warnings


def _warning_failing_elements(doc, warning):
    elements = []
    try:
        failing_ids = list(warning.GetFailingElements() or [])
    except Exception:
        failing_ids = []

    for element_id in failing_ids:
        element = doc.GetElement(element_id)
        if element is not None:
            elements.append(element)
    return elements


def _target_element_from_warning(doc, warning):
    elements = _warning_failing_elements(doc, warning)
    if not elements:
        return None, None, "No failing elements"

    # Match the Dynamo graph's List.LastItem behavior.
    element = elements[-1]
    if _is_property_line(element):
        return None, elements, "Property line"
    if _is_group_member(element):
        return None, elements, "Group member"
    return element, elements, None


def _curve_angle_degrees(curve):
    if curve is None:
        return None
    if not isinstance(curve, DB.Line):
        return None
    try:
        direction = curve.Direction
    except Exception:
        return None
    if direction is None:
        return None
    if abs(direction.X) < 1e-9 and abs(direction.Y) < 1e-9:
        return None
    return _normalize_angle_degrees(math.degrees(math.atan2(direction.Y, direction.X)))


def _rotate_curve_about_start(curve, delta_radians):
    start = curve.GetEndPoint(0)
    axis = DB.Line.CreateBound(start, start.Add(DB.XYZ.BasisZ))
    transform = DB.Transform.CreateRotationAtPoint(DB.XYZ.BasisZ, delta_radians, start)
    return curve.CreateTransformed(transform), axis, start


def _build_plan_for_warning(doc, warning, index):
    element, failing_elements, reason = _target_element_from_warning(doc, warning)
    if element is None:
        return None, reason

    try:
        location = element.Location
    except Exception:
        location = None

    if not isinstance(location, DB.LocationCurve):
        return None, "No location curve"

    try:
        curve = location.Curve
    except Exception:
        curve = None
    if curve is None:
        return None, "Missing curve"

    current_angle = _curve_angle_degrees(curve)
    if current_angle is None:
        return None, "Non-linear or vertical curve"

    target_angle = round(current_angle, ANGLE_PRECISION)
    target_angle = _normalize_angle_degrees(target_angle)
    delta_degrees = _shortest_delta_degrees(current_angle, target_angle)
    if abs(delta_degrees) <= ANGLE_EPSILON_DEGREES:
        return None, "Already rounded"

    try:
        rotated_curve, axis, origin = _rotate_curve_about_start(curve, math.radians(delta_degrees))
    except Exception as exc:
        return None, "Rotation failed: {}".format(exc)

    return {
        "warning_index": index,
        "warning_text": _warning_description(warning),
        "element": element,
        "label": _element_label(element),
        "context_text": _warning_context_text(doc, failing_elements, element),
        "context_element": _warning_context_element(failing_elements, element),
        "current_curve": curve,
        "rotated_curve": rotated_curve,
        "axis": axis,
        "origin": origin,
        "current_angle": current_angle,
        "target_angle": target_angle,
        "delta_degrees": delta_degrees,
        "delta_radians": math.radians(delta_degrees),
        "was_pinned": bool(getattr(element, "Pinned", False)),
    }, None


def _build_plans(doc, warnings):
    plans = []
    skipped = {}
    skipped_details = {}

    for index, warning in enumerate(warnings, start=1):
        plan, reason = _build_plan_for_warning(doc, warning, index)
        if plan is not None:
            plans.append(plan)
        else:
            skipped[reason] = skipped.get(reason, 0) + 1
            if reason == "Group member":
                element, failing_elements, _ = _target_element_from_warning(doc, warning)
                grouped_element = None
                if element is not None:
                    grouped_element = element
                elif failing_elements:
                    grouped_element = failing_elements[-1]
                detail = {
                    "element": grouped_element,
                    "label": _element_label(grouped_element) if grouped_element is not None else "Unknown grouped element",
                    "context_text": _warning_context_text(doc, failing_elements or [], grouped_element),
                    "context_element": _warning_context_element(failing_elements or [], grouped_element),
                    "group_element": _group_element(doc, grouped_element),
                }
                skipped_details.setdefault(reason, []).append(detail)

    plans.sort(key=lambda item: abs(item["delta_degrees"]), reverse=True)
    return plans, skipped, skipped_details


def _build_preview_report(warning_count, plans, skipped):
    lines = [
        TITLE,
        "",
        "Matched warnings: {}".format(warning_count),
        "Sketch/location lines to update: {}".format(len(plans)),
        "Angle precision: {} decimal places".format(ANGLE_PRECISION),
        "",
        "Skipped:",
    ]

    if skipped:
        for reason in sorted(skipped.keys()):
            lines.append("- {}: {}".format(reason, skipped[reason]))
    else:
        lines.append("- None")

    if plans:
        lines.extend(["", "Planned changes:"])
        for plan in plans[:MAX_DETAIL_LINES]:
            lines.append(
                "- {} | {} -> {} deg | delta {} deg".format(
                    plan["label"] + (" | Context: {}".format(plan["context_text"]) if plan.get("context_text") else ""),
                    "{:.6f}".format(plan["current_angle"]),
                    "{:.2f}".format(plan["target_angle"]),
                    "{:+.6f}".format(plan["delta_degrees"]),
                )
            )
        if len(plans) > MAX_DETAIL_LINES:
            lines.append("- ... {} more line(s)".format(len(plans) - MAX_DETAIL_LINES))

    return "\n".join(lines)


def _apply_plans(plans):
    results = {
        "updated": [],
        "failed": [],
    }

    with revit.Transaction(TITLE):
        for plan in plans:
            element = plan["element"]
            repin = False
            try:
                if plan["was_pinned"]:
                    try:
                        element.Pinned = False
                        repin = True
                    except Exception:
                        # Many sketch/profile lines expose Pinned but cannot actually
                        # be pinned or unpinned. Keep going and try to rewrite the
                        # curve anyway.
                        repin = False

                location = element.Location
                if not isinstance(location, DB.LocationCurve):
                    raise Exception("Location curve unavailable during apply")

                location.Curve = plan["rotated_curve"]
                results["updated"].append(plan)
            except Exception as exc:
                results["failed"].append({
                    "element": element,
                    "context_element": plan.get("context_element"),
                    "label": plan["label"],
                    "error": str(exc),
                })
            finally:
                if repin:
                    try:
                        element.Pinned = True
                    except Exception:
                        pass

    return results


def _build_result_report(warning_count, skipped, results):
    lines = [
        TITLE + " Results",
        "",
        "Matched warnings: {}".format(warning_count),
        "Updated: {}".format(len(results["updated"])),
        "Failed: {}".format(len(results["failed"])),
        "",
        "Skipped:",
    ]

    if skipped:
        for reason in sorted(skipped.keys()):
            lines.append("- {}: {}".format(reason, skipped[reason]))
    else:
        lines.append("- None")

    if results["updated"]:
        lines.extend(["", "Updated lines:"])
        for plan in results["updated"][:MAX_DETAIL_LINES]:
            lines.append(
                "- {} | {} -> {} deg".format(
                    plan["label"] + (" | Context: {}".format(plan["context_text"]) if plan.get("context_text") else ""),
                    "{:.6f}".format(plan["current_angle"]),
                    "{:.2f}".format(plan["target_angle"]),
                )
            )
        if len(results["updated"]) > MAX_DETAIL_LINES:
            lines.append("- ... {} more line(s)".format(len(results["updated"]) - MAX_DETAIL_LINES))

    if results["failed"]:
        lines.extend(["", "Failures:"])
        for failure in results["failed"][:MAX_DETAIL_LINES]:
            lines.append("- {} | {}".format(failure["label"], failure["error"]))
        if len(results["failed"]) > MAX_DETAIL_LINES:
            lines.append("- ... {} more failure(s)".format(len(results["failed"]) - MAX_DETAIL_LINES))

    return "\n".join(lines)


def _default_log_directory(doc):
    try:
        if doc and doc.PathName:
            folder = os.path.dirname(doc.PathName)
            if folder and os.path.isdir(folder):
                return folder
    except Exception:
        pass
    try:
        temp_dir = os.environ.get("TEMP") or os.environ.get("TMP")
        if temp_dir and os.path.isdir(temp_dir):
            return temp_dir
    except Exception:
        pass
    return os.path.expanduser("~")


def _log_file_path(doc):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(_default_log_directory(doc), "RoundOffAxisSketchLines_{}.log".format(timestamp))


def _write_log(doc, text):
    path = _log_file_path(doc)
    with open(path, "w") as log_file:
        log_file.write(text)
    return path


def _to_element_id_list(element):
    ids = List[DB.ElementId]()
    if element is not None:
        ids.Add(element.Id)
    return ids


def _preferred_ui_element(record):
    if record is None:
        return None
    return _top_parent_element(
        revit.doc,
        primary_element=record.get("element"),
        context_element=record.get("context_element"),
        explicit_group=record.get("group_element"),
    )


def _preferred_view_element(record):
    if record is None:
        return None
    return record.get("context_element") or record.get("element")


def _select_element(element):
    uidoc = revit.uidoc
    if uidoc is None or element is None:
        return False
    try:
        uidoc.Selection.SetElementIds(_to_element_id_list(element))
        return True
    except Exception:
        return False


def _zoom_to_element(element):
    uidoc = revit.uidoc
    if uidoc is None or element is None:
        return False
    try:
        uidoc.ShowElements(element.Id)
        return True
    except Exception:
        try:
            uidoc.ShowElements(_to_element_id_list(element))
            return True
        except Exception:
            return False


def _go_to_view_for_element(doc, element):
    uidoc = revit.uidoc
    if uidoc is None or doc is None or element is None:
        return False
    target_view = _owner_view(doc, element)
    if target_view is None:
        target_view = _same_level_view(doc, element)
    if target_view is None:
        return False
    try:
        uidoc.RequestViewChange(target_view)
        return True
    except Exception:
        try:
            uidoc.ActiveView = target_view
            return True
        except Exception:
            return False


def _post_edit_boundary_command():
    uidoc = revit.uidoc
    if uidoc is None:
        return False
    uiapp = uidoc.Application
    candidate_names = ["EditBoundary", "EditProfile", "EditSketch"]
    for command_name in candidate_names:
        try:
            postable = getattr(UI.PostableCommand, command_name, None)
            if postable is None:
                continue
            command_id = UI.RevitCommandId.LookupPostableCommandId(postable)
            if command_id is None:
                continue
            uiapp.PostCommand(command_id)
            return True
        except Exception:
            continue
    return False


def _edit_boundary_for_record(doc, record):
    target_element = _preferred_ui_element(record)
    if target_element is None:
        return False, "No host/context element is available for this row."
    _select_element(target_element)
    _go_to_view_for_element(doc, _preferred_view_element(record))
    if _post_edit_boundary_command():
        return True, ""
    return False, "Revit did not expose a usable sketch-edit command for this item in the current context."


def _isolate_element(element):
    uidoc = revit.uidoc
    if uidoc is None or element is None:
        return False
    try:
        active_view = uidoc.ActiveView
        if active_view is None:
            return False
        active_view.IsolateElementsTemporary(_to_element_id_list(element))
        return True
    except Exception:
        return False


class ResultBrowserWindow(object):
    def __init__(self, warning_count, skipped, skipped_details, results, log_path):
        self.warning_count = warning_count
        self.skipped = skipped
        self.skipped_details = skipped_details or {}
        self.results = results
        self.log_path = log_path
        self.items = []

        ensure_wpfui_theme()
        xaml_path = os.path.join(SCRIPT_DIR, "RoundAnglesResultsWindow.xaml")
        if not os.path.isfile(xaml_path):
            raise Exception("Missing dialog XAML: {}".format(xaml_path))

        xaml_text = File.ReadAllText(xaml_path)
        self.window = XamlReader.Parse(xaml_text)
        self.window.Title = WINDOW_TITLE + " Results"
        apply_window_title(self.window, WINDOW_TITLE + " Results")
        _set_owner(self.window)

        self._header_title = self.window.FindName("HeaderTitle")
        self._header_subtitle = self.window.FindName("HeaderSubtitle")
        self._txt_summary = self.window.FindName("TxtSummary")
        self._txt_skipped = self.window.FindName("TxtSkipped")
        self._results_list = self.window.FindName("ResultsList")
        self._txt_detail = self.window.FindName("TxtDetail")
        self._txt_log_path = self.window.FindName("TxtLogPath")
        self._btn_select = self.window.FindName("BtnSelect")
        self._btn_go_to_view = self.window.FindName("BtnGoToView")
        self._btn_edit_boundary = self.window.FindName("BtnEditBoundary")
        self._btn_zoom = self.window.FindName("BtnZoom")
        self._btn_isolate = self.window.FindName("BtnIsolate")
        self._btn_apply = self.window.FindName("BtnApply")
        self._btn_close = self.window.FindName("BtnClose")

        self._header_title.Text = WINDOW_TITLE + " Results"
        self._header_subtitle.Text = "Click a row to highlight the element in Revit. Use Select, Zoom To, or Isolate for the current row."
        self._txt_summary.Text = "Matched warnings: {}    Updated: {}    Failed: {}".format(
            warning_count,
            len(results["updated"]),
            len(results["failed"]),
        )
        self._txt_skipped.Text = self._format_skipped()
        self._txt_log_path.Text = log_path or ""

        self._build_items()
        self._populate_results()

        self._results_list.SelectionChanged += EventHandler(self._on_selection_changed)
        self._btn_select.Click += EventHandler(self._on_select_clicked)
        self._btn_go_to_view.Click += EventHandler(self._on_go_to_view_clicked)
        self._btn_edit_boundary.Click += EventHandler(self._on_edit_boundary_clicked)
        self._btn_zoom.Click += EventHandler(self._on_zoom_clicked)
        self._btn_isolate.Click += EventHandler(self._on_isolate_clicked)
        self._btn_close.Click += EventHandler(self._on_close_clicked)
        self._btn_apply.Visibility = Visibility.Collapsed
        self._update_buttons()

    def ShowDialog(self):
        return self.window.ShowDialog()

    def _format_skipped(self):
        if not self.skipped:
            return "Skipped: None"
        parts = []
        for reason in sorted(self.skipped.keys()):
            parts.append("{}: {}".format(reason, self.skipped[reason]))
        return "Skipped: " + " | ".join(parts)

    def _build_items(self):
        for plan in self.results["updated"]:
            context_suffix = " | Context: {}".format(plan["context_text"]) if plan.get("context_text") else ""
            detail_text = "Element: {}\n{}\nCurrent angle: {:.6f}\nTarget angle: {:.2f}\nDelta: {:+.6f} deg".format(
                plan["label"],
                "Context: {}".format(plan["context_text"] or "(not available)"),
                plan["current_angle"],
                plan["target_angle"],
                plan["delta_degrees"],
            )
            self.items.append({
                "status": "Updated",
                "element": plan["element"],
                "context_element": plan.get("context_element"),
                "group_element": plan.get("group_element"),
                "title": "{} | {} -> {} deg".format(
                    plan["label"] + context_suffix,
                    "{:.6f}".format(plan["current_angle"]),
                    "{:.2f}".format(plan["target_angle"]),
                ),
                "detail": _append_group_to_detail(detail_text, _resolved_group_element({
                    "element": plan["element"],
                    "context_element": plan.get("context_element"),
                    "group_element": plan.get("group_element"),
                })),
            })

        for failure in self.results["failed"]:
            detail_text = "Element: {}\nError: {}".format(failure["label"], failure["error"])
            self.items.append({
                "status": "Failed",
                "element": failure.get("element"),
                "context_element": failure.get("context_element"),
                "group_element": failure.get("group_element"),
                "title": "{} | {}".format(failure["label"], failure["error"]),
                "detail": _append_group_to_detail(detail_text, _resolved_group_element({
                    "element": failure.get("element"),
                    "context_element": failure.get("context_element"),
                    "group_element": failure.get("group_element"),
                })),
            })

        for detail in self.skipped_details.get("Group member", []):
            context_suffix = " | Context: {}".format(detail["context_text"]) if detail.get("context_text") else ""
            self.items.append({
                "status": "Grouped Skip",
                "group_skip": True,
                "element": detail.get("element"),
                "context_element": detail.get("context_element"),
                "group_element": detail.get("group_element"),
                "title": "{}{}".format(detail["label"], context_suffix),
                "detail": _append_group_to_detail(
                    "Element: {}\nContext: {}\nReason: Group member. Open Edit Group manually, then rerun the tool inside the group edit session.".format(
                        detail["label"],
                        detail.get("context_text") or "(not available)",
                    ),
                    detail.get("group_element"),
                ),
            })

    def _populate_results(self):
        self._results_list.Items.Clear()
        for item in self.items:
            self._results_list.Items.Add("[{}] {}".format(item["status"], item["title"]))
        if self._results_list.Items.Count > 0:
            self._results_list.SelectedIndex = 0

    def _current_record(self):
        try:
            index = int(self._results_list.SelectedIndex)
        except Exception:
            index = -1
        if index < 0 or index >= len(self.items):
            return None
        return self.items[index]

    def _update_buttons(self):
        record = self._current_record()
        enabled = _preferred_ui_element(record) is not None
        self._btn_select.IsEnabled = enabled
        self._btn_go_to_view.IsEnabled = enabled
        self._btn_edit_boundary.IsEnabled = enabled
        self._btn_zoom.IsEnabled = enabled
        self._btn_isolate.IsEnabled = enabled

    def _highlight_selected(self):
        record = self._current_record()
        if record is None:
            self._txt_detail.Text = ""
            self._update_buttons()
            return
        self._txt_detail.Text = record["detail"]
        _select_element(_preferred_ui_element(record))
        self._update_buttons()

    def _on_selection_changed(self, sender, args):
        self._highlight_selected()

    def _on_select_clicked(self, sender, args):
        record = self._current_record()
        if record is not None:
            _select_element(_preferred_ui_element(record))

    def _on_go_to_view_clicked(self, sender, args):
        record = self._current_record()
        if record is None:
            return
        if not _go_to_view_for_element(revit.doc, _preferred_view_element(record)):
            MessageBox.Show("No owner view was available for this item.", WINDOW_TITLE, MessageBoxButton.OK)

    def _on_edit_boundary_clicked(self, sender, args):
        record = self._current_record()
        if record is None:
            return
        ok, message = _edit_boundary_for_record(revit.doc, record)
        if not ok:
            MessageBox.Show(message, WINDOW_TITLE, MessageBoxButton.OK)

    def _on_zoom_clicked(self, sender, args):
        record = self._current_record()
        if record is None:
            return
        if _zoom_to_element(_preferred_ui_element(record)):
            return
        MessageBox.Show("Could not find a usable view for this item. Try Select and inspect the parent/context element shown in the row.", WINDOW_TITLE, MessageBoxButton.OK)

    def _on_isolate_clicked(self, sender, args):
        record = self._current_record()
        if record is None:
            return
        if _isolate_element(_preferred_ui_element(record)):
            return
        MessageBox.Show("Could not isolate this item in the active view.", WINDOW_TITLE, MessageBoxButton.OK)

    def _on_close_clicked(self, sender, args):
        self.window.DialogResult = True
        self.window.Close()


class PreviewBrowserWindow(object):
    def __init__(self, warning_count, skipped, skipped_details, plans):
        self.warning_count = warning_count
        self.skipped = skipped
        self.skipped_details = skipped_details or {}
        self.all_plans = list(plans or [])
        self.plans = list(plans or [])
        self.items = []
        self.apply_requested = False

        ensure_wpfui_theme()
        xaml_path = os.path.join(SCRIPT_DIR, "RoundAnglesResultsWindow.xaml")
        if not os.path.isfile(xaml_path):
            raise Exception("Missing dialog XAML: {}".format(xaml_path))

        xaml_text = File.ReadAllText(xaml_path)
        self.window = XamlReader.Parse(xaml_text)
        self.window.Title = WINDOW_TITLE + " Preview"
        apply_window_title(self.window, WINDOW_TITLE + " Preview")
        _set_owner(self.window)

        self._header_title = self.window.FindName("HeaderTitle")
        self._header_subtitle = self.window.FindName("HeaderSubtitle")
        self._txt_summary = self.window.FindName("TxtSummary")
        self._txt_skipped = self.window.FindName("TxtSkipped")
        self._results_list = self.window.FindName("ResultsList")
        self._selection_toolbar = self.window.FindName("SelectionToolbar")
        self._btn_remove_selected = self.window.FindName("BtnRemoveSelected")
        self._btn_reset_list = self.window.FindName("BtnResetList")
        self._txt_detail = self.window.FindName("TxtDetail")
        self._txt_log_path = self.window.FindName("TxtLogPath")
        self._btn_select = self.window.FindName("BtnSelect")
        self._btn_go_to_view = self.window.FindName("BtnGoToView")
        self._btn_edit_boundary = self.window.FindName("BtnEditBoundary")
        self._btn_zoom = self.window.FindName("BtnZoom")
        self._btn_isolate = self.window.FindName("BtnIsolate")
        self._btn_apply = self.window.FindName("BtnApply")
        self._btn_close = self.window.FindName("BtnClose")

        self._header_title.Text = WINDOW_TITLE + " Preview"
        self._header_subtitle.Text = "Review pending fixes before applying. Remove rows you do not want to process, then click Apply."
        self._txt_skipped.Text = self._format_skipped()
        self._txt_log_path.Text = "(log file is written after apply)"
        self._btn_apply.Visibility = Visibility.Visible
        self._selection_toolbar.Visibility = Visibility.Visible
        self._btn_apply.Click += EventHandler(self._on_apply_clicked)
        self._btn_remove_selected.Click += EventHandler(self._on_remove_selected_clicked)
        self._btn_reset_list.Click += EventHandler(self._on_reset_list_clicked)
        self._btn_close.Content = "Cancel"

        self._build_items()
        self._populate_results()

        self._results_list.SelectionChanged += EventHandler(self._on_selection_changed)
        self._btn_select.Click += EventHandler(self._on_select_clicked)
        self._btn_go_to_view.Click += EventHandler(self._on_go_to_view_clicked)
        self._btn_edit_boundary.Click += EventHandler(self._on_edit_boundary_clicked)
        self._btn_zoom.Click += EventHandler(self._on_zoom_clicked)
        self._btn_isolate.Click += EventHandler(self._on_isolate_clicked)
        self._btn_close.Click += EventHandler(self._on_close_clicked)
        self._update_buttons()

    def ShowDialog(self):
        return self.window.ShowDialog()

    def _format_skipped(self):
        if not self.skipped:
            return "Skipped: None"
        parts = []
        for reason in sorted(self.skipped.keys()):
            parts.append("{}: {}".format(reason, self.skipped[reason]))
        return "Skipped: " + " | ".join(parts)

    def _build_items(self):
        self.items = []
        for plan in self.plans:
            context_suffix = " | Context: {}".format(plan["context_text"]) if plan.get("context_text") else ""
            detail_text = "Element: {}\n{}\nCurrent angle: {:.6f}\nTarget angle: {:.2f}\nDelta: {:+.6f} deg".format(
                plan["label"],
                "Context: {}".format(plan["context_text"] or "(not available)"),
                plan["current_angle"],
                plan["target_angle"],
                plan["delta_degrees"],
            )
            self.items.append({
                "plan": plan,
                "element": plan["element"],
                "context_element": plan.get("context_element"),
                "group_element": plan.get("group_element"),
                "title": "{} | {} -> {} deg".format(
                    plan["label"] + context_suffix,
                    "{:.6f}".format(plan["current_angle"]),
                    "{:.2f}".format(plan["target_angle"]),
                ),
                "detail": _append_group_to_detail(detail_text, _resolved_group_element({
                    "element": plan["element"],
                    "context_element": plan.get("context_element"),
                    "group_element": plan.get("group_element"),
                })),
            })

        for detail in self.skipped_details.get("Group member", []):
            context_suffix = " | Context: {}".format(detail["context_text"]) if detail.get("context_text") else ""
            self.items.append({
                "group_skip": True,
                "element": detail.get("element"),
                "context_element": detail.get("context_element"),
                "group_element": detail.get("group_element"),
                "title": "[Grouped Skip] {}{}".format(detail["label"], context_suffix),
                "detail": _append_group_to_detail(
                    "Element: {}\nContext: {}\nReason: Group member. Open Edit Group manually, then rerun the tool inside the group edit session.".format(
                        detail["label"],
                        detail.get("context_text") or "(not available)",
                    ),
                    detail.get("group_element"),
                ),
            })

    def _populate_results(self):
        self._txt_summary.Text = "Matched warnings: {}    Pending fixes: {}".format(
            self.warning_count,
            len(self.plans),
        )
        self._results_list.Items.Clear()
        for item in self.items:
            self._results_list.Items.Add(item["title"])
        if self._results_list.Items.Count > 0:
            self._results_list.SelectedIndex = 0
        else:
            self._txt_detail.Text = ""

    def _current_record(self):
        try:
            index = int(self._results_list.SelectedIndex)
        except Exception:
            index = -1
        if index < 0 or index >= len(self.items):
            return None
        return self.items[index]

    def _update_buttons(self):
        record = self._current_record()
        enabled = _preferred_ui_element(record) is not None
        self._btn_select.IsEnabled = enabled
        self._btn_go_to_view.IsEnabled = enabled
        self._btn_edit_boundary.IsEnabled = enabled
        self._btn_zoom.IsEnabled = enabled
        self._btn_isolate.IsEnabled = enabled
        self._btn_apply.IsEnabled = len(self.items) > 0
        self._btn_remove_selected.IsEnabled = self._results_list.SelectedItems.Count > 0
        self._btn_reset_list.IsEnabled = len(self.plans) != len(self.all_plans)

    def _highlight_selected(self):
        record = self._current_record()
        if record is None:
            self._txt_detail.Text = ""
            self._update_buttons()
            return
        self._txt_detail.Text = record["detail"]
        _select_element(_preferred_ui_element(record))
        self._update_buttons()

    def _on_selection_changed(self, sender, args):
        self._highlight_selected()

    def _on_select_clicked(self, sender, args):
        record = self._current_record()
        if record is not None:
            _select_element(_preferred_ui_element(record))

    def _on_go_to_view_clicked(self, sender, args):
        record = self._current_record()
        if record is None:
            return
        if not _go_to_view_for_element(revit.doc, _preferred_view_element(record)):
            MessageBox.Show("No owner view was available for this item.", WINDOW_TITLE, MessageBoxButton.OK)

    def _on_edit_boundary_clicked(self, sender, args):
        record = self._current_record()
        if record is None:
            return
        ok, message = _edit_boundary_for_record(revit.doc, record)
        if not ok:
            MessageBox.Show(message, WINDOW_TITLE, MessageBoxButton.OK)

    def _on_zoom_clicked(self, sender, args):
        record = self._current_record()
        if record is None:
            return
        if _zoom_to_element(_preferred_ui_element(record)):
            return
        MessageBox.Show("Could not find a usable view for this item. Try Select and inspect the parent/context element shown in the row.", WINDOW_TITLE, MessageBoxButton.OK)

    def _on_isolate_clicked(self, sender, args):
        record = self._current_record()
        if record is None:
            return
        if _isolate_element(_preferred_ui_element(record)):
            return
        MessageBox.Show("Could not isolate this item in the active view.", WINDOW_TITLE, MessageBoxButton.OK)

    def _on_apply_clicked(self, sender, args):
        self.apply_requested = True
        self.window.DialogResult = True
        self.window.Close()

    def _on_remove_selected_clicked(self, sender, args):
        selected_indices = sorted([int(index) for index in list(self._results_list.SelectedIndices)], reverse=True)
        if not selected_indices:
            return
        for index in selected_indices:
            if 0 <= index < len(self.items) and self.items[index].get("group_skip"):
                continue
            if 0 <= index < len(self.plans):
                del self.plans[index]
        self._build_items()
        self._populate_results()
        self._update_buttons()

    def _on_reset_list_clicked(self, sender, args):
        self.plans = list(self.all_plans)
        self._build_items()
        self._populate_results()
        self._update_buttons()

    def _on_close_clicked(self, sender, args):
        self.apply_requested = False
        self.window.DialogResult = False
        self.window.Close()


def main():
    doc = revit.doc
    if doc is None:
        UI.TaskDialog.Show(TITLE, "No active Revit document found.")
        return

    try:
        warnings = _collect_target_warnings(doc)
        if not warnings:
            ui.uiUtils_alert(
                "No '{}' warnings were found in the current model.".format(TARGET_WARNING_TEXT),
                title=TITLE,
            )
            return

        plans, skipped, skipped_details = _build_plans(doc, warnings)
        if not plans:
            preview_text = _build_preview_report(len(warnings), plans, skipped)
            ui.uiUtils_show_text_report(
                TITLE + " Preview",
                preview_text,
                ok_text="Close",
                cancel_text=None,
                width=900,
                height=680,
            )
            return

        preview_browser = PreviewBrowserWindow(len(warnings), skipped, skipped_details, plans)
        preview_browser.ShowDialog()
        if not preview_browser.apply_requested:
            return

        results = _apply_plans(preview_browser.plans)
        result_text = _build_result_report(len(warnings), skipped, results)
        log_path = _write_log(doc, result_text)
        browser = ResultBrowserWindow(len(warnings), skipped, skipped_details, results, log_path)
        browser.ShowDialog()
    except Exception as exc:
        UI.TaskDialog.Show(
            TITLE + " - Error",
            "{}\n\n{}".format(exc, traceback.format_exc()),
        )


if __name__ == "__main__":
    main()
