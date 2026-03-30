#!python3
"""Bulk Rename – combined rename tool for views, sheets, types, materials and more."""
import os
import sys
import traceback

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.IO import File
from System.Windows.Markup import XamlReader
from System.Windows import Thickness, HorizontalAlignment, VerticalAlignment, GridLength, GridUnitType
from System.Windows.Controls import Grid as WpfGrid, CheckBox, TextBlock, ColumnDefinition
from System.Windows.Media import SolidColorBrush, Color as WpfColor

from pyrevit import DB, revit

doc = revit.doc

script_dir = os.path.dirname(__file__)
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from WWP_settings import get_tool_settings

# ─────────────────────────────────────────────────────────────────────────────
# Collectors
# ─────────────────────────────────────────────────────────────────────────────

def _skip_view_types():
    return (DB.ViewType.DrawingSheet, DB.ViewType.Schedule,
            DB.ViewType.PanelSchedule, DB.ViewType.Undefined)


def _collect_views():
    skip = _skip_view_types()
    return [v for v in DB.FilteredElementCollector(doc).OfClass(DB.View)
            if not v.IsTemplate and v.ViewType not in skip]


def _collect_sheets():
    return list(DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet))


def _collect_schedules():
    return [v for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)
            if not v.IsTemplate]


def _collect_view_templates():
    return [v for v in DB.FilteredElementCollector(doc).OfClass(DB.View)
            if v.IsTemplate]


def _collect_levels():
    return list(DB.FilteredElementCollector(doc).OfClass(DB.Level))


def _collect_grids():
    return list(DB.FilteredElementCollector(doc).OfClass(DB.Grid))


def _collect_materials():
    return list(DB.FilteredElementCollector(doc).OfClass(DB.Material))


def _collect_families():
    return list(DB.FilteredElementCollector(doc).OfClass(DB.Family))


def _collect_by_cat(bic):
    return list(
        DB.FilteredElementCollector(doc)
        .OfCategory(bic)
        .WhereElementIsElementType()
    )


def _collect_wall_types():
    return list(DB.FilteredElementCollector(doc).OfClass(DB.WallType))


def _collect_floor_types():
    return list(DB.FilteredElementCollector(doc).OfClass(DB.FloorType))


def _collect_ceiling_types():
    try:
        return list(DB.FilteredElementCollector(doc).OfClass(DB.CeilingType))
    except Exception:
        return _collect_by_cat(DB.BuiltInCategory.OST_Ceilings)


def _collect_roof_types():
    try:
        return list(DB.FilteredElementCollector(doc).OfClass(DB.RoofType))
    except Exception:
        return _collect_by_cat(DB.BuiltInCategory.OST_Roofs)


def _collect_stair_types():
    return _collect_by_cat(DB.BuiltInCategory.OST_Stairs)


def _collect_railing_types():
    return _collect_by_cat(DB.BuiltInCategory.OST_StairsRailing)


def _collect_text_types():
    return list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType))


def _collect_dimension_types():
    return list(DB.FilteredElementCollector(doc).OfClass(DB.DimensionType))


def _collect_fill_patterns():
    return list(DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement))


def _collect_line_patterns():
    return list(DB.FilteredElementCollector(doc).OfClass(DB.LinePatternElement))


def _collect_rooms():
    return list(
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Rooms)
        .WhereElementIsNotElementType()
    )


def _collect_areas():
    return list(
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Areas)
        .WhereElementIsNotElementType()
    )


def _collect_param_filters():
    return list(DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement))


# ─────────────────────────────────────────────────────────────────────────────
# Name accessors
# ─────────────────────────────────────────────────────────────────────────────

def _get_default(elem):
    return elem.Name or ""


def _set_default(elem, name):
    elem.Name = name


def _get_fill_pattern(elem):
    try:
        return elem.GetFillPattern().Name or ""
    except Exception:
        return elem.Name or ""


def _set_fill_pattern(elem, name):
    try:
        pat = elem.GetFillPattern()
        pat.Name = name
        elem.SetFillPattern(pat)
    except Exception:
        elem.Name = name


def _get_line_pattern(elem):
    try:
        return elem.GetLinePattern().Name or ""
    except Exception:
        return elem.Name or ""


def _set_line_pattern(elem, name):
    try:
        pat = elem.GetLinePattern()
        pat.Name = name
        elem.SetLinePattern(pat)
    except Exception:
        elem.Name = name


# ─────────────────────────────────────────────────────────────────────────────
# Category definitions
# ─────────────────────────────────────────────────────────────────────────────

_G = _get_default
_S = _set_default

CATEGORY_DEFS = [
    {"title": "Views",                    "collect": _collect_views,                                                              "get": _G, "set": _S},
    {"title": "Sheets",                   "collect": _collect_sheets,                                                             "get": _G, "set": _S},
    {"title": "Schedules",                "collect": _collect_schedules,                                                          "get": _G, "set": _S},
    {"title": "View Templates",           "collect": _collect_view_templates,                                                     "get": _G, "set": _S},
    {"title": "Levels",                   "collect": _collect_levels,                                                             "get": _G, "set": _S},
    {"title": "Grids",                    "collect": _collect_grids,                                                              "get": _G, "set": _S},
    {"title": "Materials",                "collect": _collect_materials,                                                          "get": _G, "set": _S},
    {"title": "Families",                 "collect": _collect_families,                                                           "get": _G, "set": _S},
    {"title": "Wall Types",               "collect": _collect_wall_types,                                                         "get": _G, "set": _S},
    {"title": "Floor Types",              "collect": _collect_floor_types,                                                        "get": _G, "set": _S},
    {"title": "Ceiling Types",            "collect": _collect_ceiling_types,                                                      "get": _G, "set": _S},
    {"title": "Roof Types",               "collect": _collect_roof_types,                                                         "get": _G, "set": _S},
    {"title": "Door Types",               "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_Doors),                      "get": _G, "set": _S},
    {"title": "Window Types",             "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_Windows),                    "get": _G, "set": _S},
    {"title": "Stair Types",              "collect": _collect_stair_types,                                                        "get": _G, "set": _S},
    {"title": "Railing Types",            "collect": _collect_railing_types,                                                      "get": _G, "set": _S},
    {"title": "Column Types (Arch)",      "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_Columns),                    "get": _G, "set": _S},
    {"title": "Column Types (Struct)",    "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_StructuralColumns),          "get": _G, "set": _S},
    {"title": "Casework Types",           "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_Casework),                   "get": _G, "set": _S},
    {"title": "Furniture Types",          "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_Furniture),                  "get": _G, "set": _S},
    {"title": "Furniture System Types",   "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_FurnitureSystems),           "get": _G, "set": _S},
    {"title": "Generic Model Types",      "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_GenericModel),               "get": _G, "set": _S},
    {"title": "Specialty Equip. Types",   "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_SpecialityEquipment),        "get": _G, "set": _S},
    {"title": "Plumbing Fixture Types",   "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_PlumbingFixtures),           "get": _G, "set": _S},
    {"title": "Mech. Equipment Types",    "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_MechanicalEquipment),        "get": _G, "set": _S},
    {"title": "Elec. Equipment Types",    "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_ElectricalEquipment),        "get": _G, "set": _S},
    {"title": "Elec. Fixture Types",      "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_ElectricalFixtures),         "get": _G, "set": _S},
    {"title": "Lighting Fixture Types",   "collect": lambda: _collect_by_cat(DB.BuiltInCategory.OST_LightingFixtures),           "get": _G, "set": _S},
    {"title": "Text Types",               "collect": _collect_text_types,                                                         "get": _G, "set": _S},
    {"title": "Dimension Types",          "collect": _collect_dimension_types,                                                    "get": _G, "set": _S},
    {"title": "Fill Patterns",            "collect": _collect_fill_patterns,                                                      "get": _get_fill_pattern, "set": _set_fill_pattern},
    {"title": "Line Patterns",            "collect": _collect_line_patterns,                                                      "get": _get_line_pattern, "set": _set_line_pattern},
    {"title": "Rooms",                    "collect": _collect_rooms,                                                              "get": _G, "set": _S},
    {"title": "Areas",                    "collect": _collect_areas,                                                              "get": _G, "set": _S},
    {"title": "Parameter Filters",        "collect": _collect_param_filters,                                                      "get": _G, "set": _S},
]

MAX_VISIBLE_ROWS = 2000

# Brush constants
_BRUSH_CHANGED = SolidColorBrush(WpfColor.FromRgb(0, 140, 0))
_BRUSH_UNCHANGED = SolidColorBrush(WpfColor.FromRgb(180, 180, 180))


# ─────────────────────────────────────────────────────────────────────────────
# Row factory
# ─────────────────────────────────────────────────────────────────────────────

def _make_row(current_name, new_name):
    """Return (Grid, CheckBox, preview_TextBlock)."""
    grid = WpfGrid()
    grid.Height = 26
    grid.Margin = Thickness(2, 0, 2, 0)

    c0 = ColumnDefinition(); c0.Width = GridLength(36)
    c1 = ColumnDefinition(); c1.Width = GridLength(1, GridUnitType.Star)
    c2 = ColumnDefinition(); c2.Width = GridLength(1, GridUnitType.Star)
    grid.ColumnDefinitions.Add(c0)
    grid.ColumnDefinitions.Add(c1)
    grid.ColumnDefinitions.Add(c2)

    cb = CheckBox()
    cb.IsChecked = True
    cb.VerticalAlignment = VerticalAlignment.Center
    cb.Margin = Thickness(8, 0, 0, 0)
    WpfGrid.SetColumn(cb, 0)

    tb_name = TextBlock()
    tb_name.Text = current_name
    tb_name.VerticalAlignment = VerticalAlignment.Center
    tb_name.Margin = Thickness(4, 0, 8, 0)
    WpfGrid.SetColumn(tb_name, 1)

    changed = new_name != current_name
    tb_preview = TextBlock()
    tb_preview.Text = new_name if changed else u"\u2014"
    tb_preview.VerticalAlignment = VerticalAlignment.Center
    tb_preview.Margin = Thickness(4, 0, 4, 0)
    tb_preview.Foreground = _BRUSH_CHANGED if changed else _BRUSH_UNCHANGED
    WpfGrid.SetColumn(tb_preview, 2)

    grid.Children.Add(cb)
    grid.Children.Add(tb_name)
    grid.Children.Add(tb_preview)

    return grid, cb, tb_preview


# ─────────────────────────────────────────────────────────────────────────────
# Dialog controller
# ─────────────────────────────────────────────────────────────────────────────

class BulkRenameDialog(object):
    def __init__(self, window, settings, save_fn):
        self.window = window
        self.settings = settings
        self.save_fn = save_fn

        self.cbo_category  = window.FindName("CboCategory")
        self.txt_search    = window.FindName("TxtSearch")
        self.txt_find      = window.FindName("TxtFind")
        self.txt_replace   = window.FindName("TxtReplace")
        self.txt_prefix    = window.FindName("TxtPrefix")
        self.txt_suffix    = window.FindName("TxtSuffix")
        self.items_panel   = window.FindName("ItemsPanel")
        self.lbl_count     = window.FindName("LblCount")
        self.lbl_status    = window.FindName("LblStatus")
        self.btn_select_all   = window.FindName("BtnSelectAll")
        self.btn_deselect_all = window.FindName("BtnDeselectAll")
        self.btn_cancel    = window.FindName("BtnCancel")
        self.btn_rename    = window.FindName("BtnRename")

        # State
        self._current_def = None
        self._all_elements = []   # [(elem, name, set_fn)]
        self._rows = []           # [{"elem", "name", "set_fn", "cb", "preview_tb"}]

        # Restore settings
        self.txt_find.Text    = getattr(settings, "last_find",    None) or ""
        self.txt_replace.Text = getattr(settings, "last_replace", None) or ""
        self.txt_prefix.Text  = getattr(settings, "last_prefix",  None) or ""
        self.txt_suffix.Text  = getattr(settings, "last_suffix",  None) or ""

        # Populate categories
        for cat_def in CATEGORY_DEFS:
            self.cbo_category.Items.Add(cat_def["title"])

        # Wire events
        self.cbo_category.SelectionChanged  += self._on_category_changed
        self.txt_search.TextChanged         += self._on_filter_changed
        self.txt_find.TextChanged           += self._on_preview_changed
        self.txt_replace.TextChanged        += self._on_preview_changed
        self.txt_prefix.TextChanged         += self._on_preview_changed
        self.txt_suffix.TextChanged         += self._on_preview_changed
        self.btn_select_all.Click           += self._on_select_all
        self.btn_deselect_all.Click         += self._on_deselect_all
        self.btn_cancel.Click               += self._on_cancel
        self.btn_rename.Click               += self._on_rename

        # Select last category (triggers load)
        last_cat = getattr(settings, "last_category", None)
        titles = [d["title"] for d in CATEGORY_DEFS]
        if last_cat and last_cat in titles:
            self.cbo_category.SelectedIndex = titles.index(last_cat)
        else:
            self.cbo_category.SelectedIndex = 0

    # ── event handlers ──────────────────────────────────────────────────────

    def _on_category_changed(self, sender, e):
        idx = self.cbo_category.SelectedIndex
        if idx < 0 or idx >= len(CATEGORY_DEFS):
            return
        self._current_def = CATEGORY_DEFS[idx]
        self._load_elements()

    def _on_filter_changed(self, sender, e):
        self._rebuild_rows()

    def _on_preview_changed(self, sender, e):
        for row in self._rows:
            new_name = self._build_new_name(row["name"])
            changed = new_name != row["name"]
            row["preview_tb"].Text = new_name if changed else u"\u2014"
            row["preview_tb"].Foreground = _BRUSH_CHANGED if changed else _BRUSH_UNCHANGED

    def _on_select_all(self, sender, e):
        for row in self._rows:
            row["cb"].IsChecked = True

    def _on_deselect_all(self, sender, e):
        for row in self._rows:
            row["cb"].IsChecked = False

    def _on_cancel(self, sender, e):
        self.window.Close()

    def _on_rename(self, sender, e):
        to_rename = []
        for row in self._rows:
            if not row["cb"].IsChecked:
                continue
            new_name = self._build_new_name(row["name"]).strip()
            if new_name and new_name != row["name"]:
                to_rename.append((row["elem"], row["name"], new_name, row["set_fn"]))

        if not to_rename:
            self.lbl_status.Text = "Nothing to rename — check selection and inputs."
            self.lbl_status.Foreground = SolidColorBrush(WpfColor.FromRgb(180, 80, 0))
            return

        # Persist settings before applying
        self._save_settings()

        t = DB.Transaction(doc, "Bulk Rename")
        t.Start()
        renamed = 0
        failed = []
        try:
            for elem, old_name, new_name, set_fn in to_rename:
                try:
                    set_fn(elem, new_name)
                    renamed += 1
                except Exception as ex:
                    failed.append(u"{} \u2192 {} ({})".format(old_name, new_name, str(ex)[:120]))
        finally:
            if t.HasStarted() and not t.HasEnded():
                t.Commit()

        # Reload to show updated names
        self._load_elements()

        if failed:
            msg = u"Renamed {}. {} failed.".format(renamed, len(failed))
            self.lbl_status.Foreground = SolidColorBrush(WpfColor.FromRgb(180, 80, 0))
        else:
            msg = u"Renamed {} item(s).".format(renamed)
            self.lbl_status.Foreground = SolidColorBrush(WpfColor.FromRgb(0, 120, 0))
        self.lbl_status.Text = msg

    # ── helpers ─────────────────────────────────────────────────────────────

    def _build_new_name(self, current):
        find    = self.txt_find.Text    or ""
        replace = self.txt_replace.Text or ""
        prefix  = self.txt_prefix.Text  or ""
        suffix  = self.txt_suffix.Text  or ""
        name = current
        if find:
            name = name.replace(find, replace)
        if prefix:
            name = prefix + name
        if suffix:
            name = name + suffix
        return name

    def _load_elements(self):
        cat_def = self._current_def
        if not cat_def:
            return
        self.lbl_count.Text = "Loading…"
        self.lbl_status.Text = ""
        try:
            elems = cat_def["collect"]()
        except Exception as ex:
            self.lbl_count.Text = ""
            self.lbl_status.Text = "Error collecting elements: {}".format(ex)
            self.lbl_status.Foreground = SolidColorBrush(WpfColor.FromRgb(180, 0, 0))
            return

        get_fn = cat_def["get"]
        set_fn = cat_def["set"]
        pairs = []
        for elem in elems:
            try:
                name = get_fn(elem)
                pairs.append((elem, name, set_fn))
            except Exception:
                pass
        pairs.sort(key=lambda t: t[1].lower())
        self._all_elements = pairs
        self._rebuild_rows()

    def _rebuild_rows(self):
        search = (self.txt_search.Text or "").lower()
        self.items_panel.Children.Clear()
        self._rows = []

        filtered = [t for t in self._all_elements
                    if not search or search in t[1].lower()]

        shown = filtered[:MAX_VISIBLE_ROWS]
        for elem, name, set_fn in shown:
            new_name = self._build_new_name(name)
            row_grid, cb, preview_tb = _make_row(name, new_name)
            self.items_panel.Children.Add(row_grid)
            self._rows.append({"elem": elem, "name": name, "set_fn": set_fn,
                                "cb": cb, "preview_tb": preview_tb})

        total = len(filtered)
        if len(shown) < total:
            self.lbl_count.Text = "Showing {} of {} item(s) (capped at {}).".format(
                len(shown), total, MAX_VISIBLE_ROWS)
        else:
            self.lbl_count.Text = "{} item(s).".format(total)

    def _save_settings(self):
        if self._current_def:
            self.settings.last_category = self._current_def["title"]
        self.settings.last_find    = self.txt_find.Text    or ""
        self.settings.last_replace = self.txt_replace.Text or ""
        self.settings.last_prefix  = self.txt_prefix.Text  or ""
        self.settings.last_suffix  = self.txt_suffix.Text  or ""
        try:
            self.save_fn()
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main():
    settings, save_fn = get_tool_settings("BulkRename", doc=doc)

    xaml_path = os.path.join(script_dir, "RenameWindow.xaml")
    window = XamlReader.Parse(File.ReadAllText(xaml_path))

    BulkRenameDialog(window, settings, save_fn)
    window.ShowDialog()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        from Autodesk.Revit.UI import TaskDialog
        TaskDialog.Show("Bulk Rename – Error", traceback.format_exc())
