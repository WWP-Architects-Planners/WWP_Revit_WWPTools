#!python3
# -*- coding: utf-8 -*-

import os
import sys
from pyrevit import revit, DB
import WWP_uiUtils as ui
from pyrevit.framework import EventHandler
from System.IO import File
from System.Windows import Visibility
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Markup import XamlReader

script_dir = os.path.dirname(__file__)
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))
if lib_path not in sys.path:
    sys.path.append(lib_path)

import WWP_colorSchemeUtils as csu


def _scheme_area_scheme_name(doc, scheme):
    try:
        area_scheme_id = getattr(scheme, "AreaSchemeId", None)
        if area_scheme_id:
            area_scheme = doc.GetElement(area_scheme_id)
            if area_scheme:
                return getattr(area_scheme, "Name", "") or ""
    except Exception:
        pass
    try:
        get_area_scheme_id = getattr(scheme, "GetAreaSchemeId", None)
        if callable(get_area_scheme_id):
            area_scheme_id = get_area_scheme_id()
            if area_scheme_id:
                area_scheme = doc.GetElement(area_scheme_id)
                if area_scheme:
                    return getattr(area_scheme, "Name", "") or ""
    except Exception:
        pass
    return ""


def _category_name(doc, category_id):
    if category_id is None:
        return "Unknown Category"
    try:
        cat = doc.GetElement(category_id)
        cat_name = getattr(cat, "Name", "") if cat else ""
        if cat_name:
            return cat_name
    except Exception:
        pass
    try:
        return DB.LabelUtils.GetLabelFor(category_id)
    except Exception:
        pass
    try:
        import System
        cat_int = int(category_id.IntegerValue)
        bic = System.Enum.ToObject(DB.BuiltInCategory, cat_int)
        label = DB.LabelUtils.GetLabelFor(bic)
        if label:
            return label
        return "Category {}".format(cat_int)
    except Exception:
        return "Unknown Category"


def _scheme_display_name(doc, scheme):
    scheme_name = getattr(scheme, "Name", "") or "Color Scheme"
    area_name = _scheme_area_scheme_name(doc, scheme)
    if area_name:
        return "Area({}):{}".format(area_name, scheme_name)
    return "{}: {}".format(_category_name(doc, scheme.CategoryId), scheme_name)


def _set_owner(window):
    try:
        helper = WindowInteropHelper(window)
        helper.Owner = revit.uidoc.Application.MainWindowHandle
    except Exception:
        pass


def _show_selection_dialog(display_names):
    xaml_path = os.path.join(script_dir, "CopyColorSchemeSetup.xaml")
    if not os.path.isfile(xaml_path):
        raise Exception("Missing dialog XAML: {}".format(xaml_path))

    xaml_text = File.ReadAllText(xaml_path)
    window = XamlReader.Parse(xaml_text)
    _set_owner(window)

    source_combo = window.FindName("SourceSchemeCombo")
    target_combo = window.FindName("TargetSchemeCombo")
    validation = window.FindName("ValidationText")
    ok_btn = window.FindName("OkButton")
    cancel_btn = window.FindName("CancelButton")

    for name in display_names:
        source_combo.Items.Add(name)
        target_combo.Items.Add(name)

    if source_combo.Items.Count > 0:
        source_combo.SelectedIndex = 0
    if target_combo.Items.Count > 1:
        target_combo.SelectedIndex = 1
    elif target_combo.Items.Count > 0:
        target_combo.SelectedIndex = 0

    result = {"ok": False}

    def _set_validation(msg):
        if msg:
            validation.Text = msg
            validation.Visibility = Visibility.Visible
        else:
            validation.Text = ""
            validation.Visibility = Visibility.Collapsed

    def _on_ok(sender, args):
        source_idx = int(source_combo.SelectedIndex)
        target_idx = int(target_combo.SelectedIndex)
        if source_idx < 0 or target_idx < 0:
            _set_validation("Please select both source and target schemes.")
            return
        if source_idx == target_idx:
            _set_validation("Source and target must be different schemes.")
            return
        result["ok"] = True
        result["source_index"] = source_idx
        result["target_index"] = target_idx
        window.DialogResult = True
        window.Close()

    def _on_cancel(sender, args):
        window.DialogResult = False
        window.Close()

    ok_btn.Click += EventHandler(_on_ok)
    cancel_btn.Click += EventHandler(_on_cancel)
    if window.ShowDialog() != True:
        return None
    return result if result.get("ok") else None


def main():
    doc = revit.doc
    schemes = csu.collect_color_fill_schemes(doc)
    if not schemes:
        ui.uiUtils_alert("No Color Fill Schemes found in this model.", title="Copy Color Scheme")
        return

    names = [_scheme_display_name(doc, s) for s in schemes]
    pick = _show_selection_dialog(names)
    if not pick:
        return

    source = schemes[int(pick["source_index"])]
    selected_target = schemes[int(pick["target_index"])]

    with revit.Transaction("Copy Color Scheme"):
        source_name = (getattr(source, "Name", "") or "Color Scheme").strip()
        target = selected_target

        # Target acts as scope reference unless names already match.
        selected_name = (getattr(selected_target, "Name", "") or "").strip()
        if selected_name.lower() != source_name.lower():
            all_schemes = csu.collect_color_fill_schemes(doc)
            existing_same_name = csu.find_scheme_in_scope_by_name(all_schemes, selected_target, source_name)
            if existing_same_name is not None:
                target = existing_same_name
            else:
                new_id = selected_target.Duplicate(source_name)
                created = doc.GetElement(new_id)
                if created is None:
                    ui.uiUtils_alert(
                        "Failed to create new scheme '{}' in selected target scope.".format(source_name),
                        title="Copy Color Scheme",
                    )
                    return
                target = created

        ok, error = csu.copy_scheme_data(source, target)

    if not ok:
        ui.uiUtils_alert("Failed to update target Color Scheme.\n\n{}".format(error or "Unknown error"), title="Copy Color Scheme")
        return

    try:
        revit.get_selection().set_to([target.Id])
    except Exception:
        pass

    ui.uiUtils_alert("Updated scheme: {}".format(getattr(target, "Name", "Color Scheme")), title="Copy Color Scheme")


if __name__ == "__main__":
    main()
