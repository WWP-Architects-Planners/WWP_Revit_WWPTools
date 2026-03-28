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
from System.Windows.Controls import ListBoxItem, ComboBoxItem
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
WINDOW_TITLE = "Convert Lines to Detail Items"


def _set_owner(window):
    helper = WindowInteropHelper(window)
    helper.Owner = uidoc.Application.MainWindowHandle


def collect_frr_lines_in_view(view):
    lines = []
    for elem in (DB.FilteredElementCollector(doc, view.Id)
                 .OfCategory(DB.BuiltInCategory.OST_Lines)
                 .WhereElementIsNotElementType()
                 .ToElements()):
        try:
            style = elem.LineStyle
            if style and "FRR" in (style.Name or ""):
                lines.append(elem)
        except Exception:
            pass
    return lines


def collect_detail_component_families():
    """Return list of (family_name, symbol) for all detail component families."""
    symbols = []
    for sym in (DB.FilteredElementCollector(doc)
                .OfClass(DB.FamilySymbol)
                .ToElements()):
        cat = sym.Category
        if cat is None:
            continue
        if cat.Id.IntegerValue == int(DB.BuiltInCategory.OST_DetailComponents):
            fam = sym.Family
            fam_name = fam.Name if fam else ""
            sym_name = DB.Element.Name.GetValue(sym)
            symbols.append(("{} : {}".format(fam_name, sym_name), sym))
    symbols.sort(key=lambda x: x[0])
    return symbols


class ConvertDialog(object):
    def __init__(self, frr_lines, family_symbols):
        xaml_path = os.path.join(script_dir, "FireRatingConvertWindow.xaml")
        xaml_text = File.ReadAllText(xaml_path)
        self.window = XamlReader.Parse(xaml_text)
        self.window.Title = WINDOW_TITLE
        apply_window_title(self.window, WINDOW_TITLE)
        _set_owner(self.window)

        self._lbl_count = self.window.FindName("LblLineCount")
        self._combo_family = self.window.FindName("CboFamily")
        self._btn_convert = self.window.FindName("BtnConvert")
        self._btn_cancel = self.window.FindName("BtnCancel")
        self._lbl_status = self.window.FindName("LblStatus")

        self._lbl_count.Text = "{} FRR line(s) found in current view.".format(len(frr_lines))

        for label, sym in family_symbols:
            item = ComboBoxItem()
            item.Content = label
            item.Tag = sym
            self._combo_family.Items.Add(item)
        if self._combo_family.Items.Count > 0:
            self._combo_family.SelectedIndex = 0

        self._btn_convert.Click += EventHandler(self._on_convert)
        self._btn_cancel.Click += EventHandler(self._on_cancel)

        self.selected_symbol = None
        self.accepted = False

    def _on_convert(self, sender, args):
        sel = self._combo_family.SelectedItem
        if sel is None:
            self._lbl_status.Text = "Select a detail component family."
            return
        self.selected_symbol = sel.Tag
        self.accepted = True
        self.window.DialogResult = True
        self.window.Close()

    def _on_cancel(self, sender, args):
        self.window.DialogResult = False
        self.window.Close()

    def ShowDialog(self):
        return self.window.ShowDialog()


def convert_lines_to_components(frr_lines, symbol, active_view):
    """Place a detail component at the midpoint of each FRR line, then delete the line."""
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()

    placed = 0
    t = DB.Transaction(doc, "Convert Lines to Detail Items")
    t.Start()
    try:
        from System.Collections.Generic import List
        ids_to_delete = List[DB.ElementId]()
        for line_elem in frr_lines:
            try:
                curve = line_elem.GeometryCurve
                mid_param = (curve.GetEndParameter(0) + curve.GetEndParameter(1)) / 2.0
                mid_pt = curve.Evaluate(mid_param, False)
                # Place component at midpoint
                doc.Create.NewFamilyInstance(
                    mid_pt, symbol, active_view
                )
                ids_to_delete.Add(line_elem.Id)
                placed += 1
            except Exception:
                pass
        if ids_to_delete.Count > 0:
            doc.Delete(ids_to_delete)
        t.Commit()
    except Exception:
        t.RollBack()
        raise
    return placed


def main():
    if doc is None:
        ui.uiUtils_alert("Open a project first.", title=WINDOW_TITLE)
        return

    active_view = doc.ActiveView
    frr_lines = collect_frr_lines_in_view(active_view)
    if not frr_lines:
        ui.uiUtils_alert(
            "No FRR detail lines found in the current view.",
            title=WINDOW_TITLE,
        )
        return

    family_symbols = collect_detail_component_families()
    if not family_symbols:
        ui.uiUtils_alert(
            "No detail component families found in the project.\n"
            "Load a detail component family first.",
            title=WINDOW_TITLE,
        )
        return

    dialog = ConvertDialog(frr_lines, family_symbols)
    if not dialog.ShowDialog():
        return

    placed = convert_lines_to_components(
        frr_lines, dialog.selected_symbol, active_view
    )
    ui.uiUtils_alert(
        "Converted {} FRR line(s) to detail components.".format(placed),
        title=WINDOW_TITLE,
    )


if __name__ == "__main__":
    try:
        main()
    except Exception:
        ui.uiUtils_alert(traceback.format_exc(), title=WINDOW_TITLE)
