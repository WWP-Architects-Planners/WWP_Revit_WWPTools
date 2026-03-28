#!python3
# -*- coding: utf-8 -*-
import os
import sys
import traceback
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from pyrevit.framework import EventHandler
from System.IO import File
from System.Windows.Markup import XamlReader
from System.Windows.Interop import WindowInteropHelper

script_dir = os.path.dirname(__file__)
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "..", "lib"))
if lib_path not in sys.path:
    sys.path.append(lib_path)

import WWP_uiUtils as ui
from WWP_versioning import apply_window_title
from pyrevit import DB

uidoc = getattr(__revit__, "ActiveUIDocument", None)
doc = uidoc.Document if uidoc else None
WINDOW_TITLE = "Create Fire Rating Lines — All Walls"

# Maps FRR rating substrings to line style name substrings
FRR_STYLE_MAP = [
    ("0HR", "0HR"),
    ("3/4HR", "3-4HR"),
    ("1HR", "1HR"),
    ("1.5HR", "1-5HR"),
    ("2HR", "2HR"),
    ("3HR", "3HR"),
    ("4HR", "4HR"),
]


def _set_owner(window):
    helper = WindowInteropHelper(window)
    helper.Owner = uidoc.Application.MainWindowHandle


def get_all_line_styles():
    lines_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
    return {sub.Name: sub.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
            for sub in lines_cat.SubCategories}


def find_style_for_frr(frr_value, all_styles):
    """Find best matching line style for a given FRR value."""
    if not frr_value:
        return None
    frr_upper = frr_value.strip().upper()
    # Try exact match first (e.g., "WWP - FRR - 1HR")
    for name, style in all_styles.items():
        if "FRR" in name.upper() and frr_upper in name.upper():
            return style
    # Try with dashes replacing slashes (e.g., 3/4HR → 3-4HR)
    frr_dash = frr_upper.replace("/", "-")
    for name, style in all_styles.items():
        if "FRR" in name.upper() and frr_dash in name.upper():
            return style
    return None


class FireRatingAllDialog(object):
    def __init__(self):
        xaml_path = os.path.join(script_dir, "FireRatingWindow.xaml")
        xaml_text = File.ReadAllText(xaml_path)
        self.window = XamlReader.Parse(xaml_text)
        self.window.Title = WINDOW_TITLE
        apply_window_title(self.window, WINDOW_TITLE)
        _set_owner(self.window)

        self._txt_param = self.window.FindName("TxtFrrParam")
        self._btn_run = self.window.FindName("BtnRun")
        self._btn_cancel = self.window.FindName("BtnCancel")
        self._lbl_status = self.window.FindName("LblStatus")
        self._lbl_count = self.window.FindName("LblWallCount")

        self._txt_param.Text = "FRR Walls"
        self._update_count()

        self._btn_run.Click += EventHandler(self._on_run)
        self._btn_cancel.Click += EventHandler(self._on_cancel)
        self._txt_param.TextChanged += EventHandler(self._on_param_changed)

        self.frr_param = None
        self.accepted = False

    def _update_count(self):
        try:
            active_view = doc.ActiveView
            count = (DB.FilteredElementCollector(doc, active_view.Id)
                     .OfCategory(DB.BuiltInCategory.OST_Walls)
                     .WhereElementIsNotElementType()
                     .GetElementCount())
            self._lbl_count.Text = "{} wall(s) in current view.".format(count)
        except Exception:
            self._lbl_count.Text = ""

    def _on_param_changed(self, sender, args):
        self._lbl_status.Text = ""

    def _on_run(self, sender, args):
        self.frr_param = (self._txt_param.Text or "").strip()
        if not self.frr_param:
            self._lbl_status.Text = "Parameter name is required."
            return
        self.accepted = True
        self.window.DialogResult = True
        self.window.Close()

    def _on_cancel(self, sender, args):
        self.window.DialogResult = False
        self.window.Close()

    def ShowDialog(self):
        return self.window.ShowDialog()


def create_frr_lines_for_walls(walls, frr_param_name, active_view):
    all_styles = get_all_line_styles()
    created = 0
    skipped_no_param = 0
    skipped_zero = 0

    t = DB.Transaction(doc, "Create Fire Rating Lines")
    t.Start()
    try:
        for wall in walls:
            loc = wall.Location
            if not isinstance(loc, DB.LocationCurve):
                continue
            curve = loc.Curve

            param = wall.LookupParameter(frr_param_name)
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

    return created, skipped_no_param, skipped_zero


def main():
    if doc is None:
        ui.uiUtils_alert("Open a project first.", title=WINDOW_TITLE)
        return

    active_view = doc.ActiveView
    if active_view is None:
        ui.uiUtils_alert("No active view.", title=WINDOW_TITLE)
        return

    dialog = FireRatingAllDialog()
    if not dialog.ShowDialog():
        return

    walls = [
        e for e in DB.FilteredElementCollector(doc, active_view.Id)
            .OfCategory(DB.BuiltInCategory.OST_Walls)
            .WhereElementIsNotElementType()
            .ToElements()
    ]

    if not walls:
        ui.uiUtils_alert("No walls found in the current view.", title=WINDOW_TITLE)
        return

    created, skip_param, skip_zero = create_frr_lines_for_walls(
        walls, dialog.frr_param, active_view
    )

    lines = ["Created {} fire rating line(s).".format(created)]
    if skip_zero:
        lines.append("{} wall(s) skipped — rated 0HR or no rating.".format(skip_zero))
    if skip_param:
        lines.append("{} wall(s) skipped — missing '{}' parameter.".format(
            skip_param, dialog.frr_param))
    ui.uiUtils_alert("\n".join(lines), title=WINDOW_TITLE)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        ui.uiUtils_alert(traceback.format_exc(), title=WINDOW_TITLE)
