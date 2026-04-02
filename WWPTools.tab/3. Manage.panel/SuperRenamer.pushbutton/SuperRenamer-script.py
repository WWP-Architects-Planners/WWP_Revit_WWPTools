import clr
import os
import sys
import traceback

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from pyrevit import DB, revit
from pyrevit.framework import EventHandler
from System.IO import File
from System.Windows import RoutedEventHandler
from System.Windows.Controls import SelectionChangedEventHandler
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Markup import XamlReader

script_dir = os.path.dirname(__file__)
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

import WWP_uiUtils as ui


# ---------------------------------------------------------------------------
# Category definitions
# ---------------------------------------------------------------------------

_SELECTION_CATEGORIES = {"Types (Selection)"}

_SOURCE_LABELS = {
    "Types (Selection)": "Source: Current selection (select elements or types first)",
}


def _source_label(category):
    return _SOURCE_LABELS.get(category, "Source: All elements in document")


# ---------------------------------------------------------------------------
# Element collection
# ---------------------------------------------------------------------------

def collect_elements(doc, category):
    fec = DB.FilteredElementCollector
    if category == "Materials":
        return list(fec(doc).OfClass(DB.Material).ToElements())

    elif category == "Views":
        return [
            v for v in fec(doc).OfClass(DB.View).ToElements()
            if not v.IsTemplate
            and v.ViewType != DB.ViewType.Schedule
            and v.ViewType != DB.ViewType.DrawingSheet
            and v.ViewType != DB.ViewType.Internal
        ]

    elif category == "View Templates":
        return [v for v in fec(doc).OfClass(DB.View).ToElements() if v.IsTemplate]

    elif category == "Sheets":
        return list(fec(doc).OfClass(DB.ViewSheet).ToElements())

    elif category == "Levels":
        return list(fec(doc).OfClass(DB.Level).WhereElementIsNotElementType().ToElements())

    elif category == "Grids":
        return list(fec(doc).OfClass(DB.Grid).WhereElementIsNotElementType().ToElements())

    elif category == "Rooms":
        return [
            e for e in fec(doc).OfCategory(DB.BuiltInCategory.OST_Rooms)
                               .WhereElementIsNotElementType()
                               .ToElements()
            if e is not None and _get_name(e)
        ]

    elif category == "Spaces":
        return [
            e for e in fec(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
                               .WhereElementIsNotElementType()
                               .ToElements()
            if e is not None
        ]

    elif category == "Areas":
        return [
            e for e in fec(doc).OfCategory(DB.BuiltInCategory.OST_Areas)
                               .WhereElementIsNotElementType()
                               .ToElements()
            if e is not None
        ]

    elif category == "View Filters":
        return list(fec(doc).OfClass(DB.FilterElement).ToElements())

    elif category == "Phases":
        try:
            return list(doc.Phases)
        except Exception:
            return []

    elif category == "Types (Selection)":
        return _get_selected_types(doc)

    return []


def _get_selected_types(doc):
    try:
        selection = revit.get_selection()
        selected = list(selection.elements) if selection else []
    except Exception:
        return []

    types = []
    seen = set()
    for element in selected:
        if isinstance(element, DB.ElementType):
            eid = element.Id.IntegerValue
            if eid not in seen:
                types.append(element)
                seen.add(eid)
            continue
        try:
            type_id = element.GetTypeId()
            if type_id == DB.ElementId.InvalidElementId:
                continue
            type_elem = doc.GetElement(type_id)
            if isinstance(type_elem, DB.ElementType):
                eid = type_elem.Id.IntegerValue
                if eid not in seen:
                    types.append(type_elem)
                    seen.add(eid)
        except Exception:
            pass

    return types


# ---------------------------------------------------------------------------
# Name helpers
# ---------------------------------------------------------------------------

def _get_name(element):
    try:
        return element.Name or ""
    except Exception:
        return ""


def _set_name(element, name):
    element.Name = name


def _build_new_name(current, find_text, replace_text, prefix, suffix):
    new_name = current
    if find_text:
        new_name = new_name.replace(find_text, replace_text)
    if prefix:
        new_name = "{}{}".format(prefix, new_name)
    if suffix:
        new_name = "{}{}".format(new_name, suffix)
    return new_name


# ---------------------------------------------------------------------------
# Plan & apply
# ---------------------------------------------------------------------------

def plan_renames(elements, find_text, replace_text, prefix, suffix):
    existing_lower = {_get_name(e).lower() for e in elements if _get_name(e)}
    planned = []
    skipped = []

    for elem in elements:
        old_name = _get_name(elem)
        new_name = _build_new_name(old_name, find_text, replace_text, prefix, suffix).strip()

        if new_name == old_name:
            continue
        if not new_name:
            skipped.append((old_name, new_name, "empty name"))
            continue
        if len(new_name) > 255:
            skipped.append((old_name, new_name, "name too long"))
            continue
        if new_name.lower() in existing_lower and new_name.lower() != old_name.lower():
            skipped.append((old_name, new_name, "name conflict"))
            continue

        planned.append((elem, old_name, new_name))
        existing_lower.discard(old_name.lower())
        existing_lower.add(new_name.lower())

    return planned, skipped


def apply_renames(doc, planned, category):
    renamed = []
    failed = []
    t = DB.Transaction(doc, "Super Rename {}".format(category))
    try:
        t.Start()
        for elem, old_name, new_name in planned:
            try:
                _set_name(elem, new_name)
                renamed.append((old_name, new_name))
            except Exception as ex:
                failed.append((old_name, new_name, str(ex)))
        t.Commit()
    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        return [], failed + [("<transaction>", "<commit>", str(ex))]

    return renamed, failed


# ---------------------------------------------------------------------------
# WPF theme loader
# ---------------------------------------------------------------------------

def _ensure_theme():
    try:
        ver = int(str(__revit__.Application.VersionNumber))
    except Exception:
        ver = None
    dll_name = (
        "WWPTools.WpfUI.net8.0-windows.dll"
        if ver and ver >= 2025
        else "WWPTools.WpfUI.net48.dll"
    )
    dll_path = os.path.join(lib_path, dll_name)
    if not os.path.isfile(dll_path):
        return
    try:
        if hasattr(clr, "AddReferenceToFileAndPath"):
            clr.AddReferenceToFileAndPath(dll_path)
        else:
            clr.AddReference(dll_path)
    except Exception:
        pass


def _set_owner(window):
    try:
        helper = WindowInteropHelper(window)
        helper.Owner = revit.uidoc.Application.MainWindowHandle
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Dialog
# ---------------------------------------------------------------------------

def show_dialog():
    _ensure_theme()

    xaml_path = os.path.join(script_dir, "SuperRenamer.xaml")
    if not os.path.isfile(xaml_path):
        raise Exception("Missing XAML file: {}".format(xaml_path))

    xaml_text = File.ReadAllText(xaml_path)
    window = XamlReader.Parse(xaml_text)
    _set_owner(window)

    cmb_category = window.FindName("CmbCategory")
    txt_source   = window.FindName("TxtSource")
    txt_find     = window.FindName("TxtFind")
    txt_replace  = window.FindName("TxtReplace")
    txt_prefix   = window.FindName("TxtPrefix")
    txt_suffix   = window.FindName("TxtSuffix")
    btn_cancel   = window.FindName("BtnCancel")
    btn_apply    = window.FindName("BtnApply")

    cmb_category.SelectedIndex = 0

    result = [None]  # use list so closure can mutate it

    def _selected_category():
        item = cmb_category.SelectedItem
        try:
            return str(item.Content or "")
        except Exception:
            return str(item or "")

    def _on_category_changed(sender, args):
        txt_source.Text = _source_label(_selected_category())

    def _on_apply(sender, args):
        result[0] = {
            "category": _selected_category(),
            "find":    txt_find.Text or "",
            "replace": txt_replace.Text or "",
            "prefix":  txt_prefix.Text or "",
            "suffix":  txt_suffix.Text or "",
        }
        window.DialogResult = True
        window.Close()

    def _on_cancel(sender, args):
        window.DialogResult = False
        window.Close()

    cmb_category.SelectionChanged += SelectionChangedEventHandler(_on_category_changed)
    btn_apply.Click  += RoutedEventHandler(_on_apply)
    btn_cancel.Click += RoutedEventHandler(_on_cancel)

    if window.ShowDialog() != True:
        return None
    return result[0]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    doc = revit.doc

    inputs = show_dialog()
    if not inputs:
        return

    category     = inputs["category"]
    find_text    = inputs["find"]
    replace_text = inputs["replace"]
    prefix       = inputs["prefix"]
    suffix       = inputs["suffix"]

    if not any([find_text, prefix, suffix]):
        ui.uiUtils_alert(
            "Provide at least a Find text, Prefix, or Suffix value.",
            title="Super Renamer",
        )
        return

    elements = collect_elements(doc, category)
    if not elements:
        if category in _SELECTION_CATEGORIES:
            msg = "No types found in the current selection.\nSelect elements or types in the Project Browser first."
        else:
            msg = "No {} found in the document.".format(category.lower())
        ui.uiUtils_alert(msg, title="Super Renamer")
        return

    planned, skipped = plan_renames(elements, find_text, replace_text, prefix, suffix)

    if not planned:
        ui.uiUtils_alert(
            "No elements matched the criteria.\nSkipped: {}".format(len(skipped)),
            title="Super Renamer",
        )
        return

    # --- Preview ---
    lines = [
        "Category:  {}".format(category),
        "To rename: {}".format(len(planned)),
        "Skipped:   {}".format(len(skipped)),
        "",
    ]
    for _, old, new in planned[:300]:
        lines.append("{}  \u2192  {}".format(old, new))
    if len(planned) > 300:
        lines.append("... and {} more".format(len(planned) - 300))
    if skipped:
        lines.append("")
        lines.append("Skipped (conflicts / invalid):")
        for old, new, reason in skipped[:50]:
            lines.append("  {}  \u2192  {}  [{}]".format(old, new, reason))

    proceed = ui.uiUtils_show_text_report(
        "Super Renamer \u2013 Preview",
        "\n".join(lines),
        ok_text="Apply",
        cancel_text="Cancel",
        width=720,
        height=520,
    )
    if not proceed:
        return

    # --- Apply ---
    renamed, failed = apply_renames(doc, planned, category)

    result_lines = [
        "Renamed: {}".format(len(renamed)),
        "Failed:  {}".format(len(failed)),
    ]
    if failed:
        result_lines += ["", "Failed (first 20):"]
        for old, new, ex in failed[:20]:
            result_lines.append("  {}  \u2192  {}  ({})".format(old, new, ex))

    ui.uiUtils_show_text_report(
        "Super Renamer \u2013 Results",
        "\n".join(result_lines),
        ok_text="Close",
        cancel_text=None,
        width=580,
        height=380,
    )


if __name__ == "__main__":
    try:
        main()
    except Exception:
        ui.uiUtils_alert(traceback.format_exc(), title="Super Renamer \u2013 Error")
