# -*- coding: utf-8 -*-
"""
WWP Manual Revisions Tool
Toggles between native Revit revision schedule and manual
4-parameter revision block on the active sheet (or selected sheets).

Parameters expected on Sheet:
  - WWP_ShowManualRevisions_YesNo  (Yes/No)
  - WWP_Rev_Left_Dates             (Text)
  - WWP_Rev_Left_Descs             (Text)
  - WWP_Rev_Right_Dates            (Text)
  - WWP_Rev_Right_Descs            (Text)
"""

__title__ = "Manual\nRevisions"
__doc__ = "Toggle and fill manual revision block on sheet(s)."
__author__ = "WWP"

import os
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from pyrevit import forms, revit, DB, script
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory

import System
from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxResult
from System.Windows.Markup import XamlReader
from System.IO import StringReader

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

PARAM_SHOW        = "WWP_ShowManualRevisions_YesNo"
PARAM_LEFT_DATES  = "WWP_Rev_Left_Dates"
PARAM_LEFT_DESCS  = "WWP_Rev_Left_Descs"
PARAM_RIGHT_DATES = "WWP_Rev_Right_Dates"
PARAM_RIGHT_DESCS = "WWP_Rev_Right_Descs"

doc   = revit.doc
uidoc = revit.uidoc

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_param(element, param_name):
    """Return the Parameter object or None."""
    p = element.LookupParameter(param_name)
    return p


def get_target_sheets():
    """
    Return sheets to operate on:
      - If sheets are pre-selected in the project browser, use those.
      - Otherwise fall back to the active view if it is a sheet.
      - Otherwise ask the user to pick.
    """
    selection_ids = uidoc.Selection.GetElementIds()
    sheets = []

    for eid in selection_ids:
        el = doc.GetElement(eid)
        if isinstance(el, DB.ViewSheet):
            sheets.append(el)

    if sheets:
        return sheets

    active = uidoc.ActiveView
    if isinstance(active, DB.ViewSheet):
        return [active]

    # Let user pick from all sheets
    all_sheets = (
        FilteredElementCollector(doc)
        .OfClass(DB.ViewSheet)
        .ToElements()
    )
    sheet_list = sorted(all_sheets, key=lambda s: s.SheetNumber)

    picked = forms.SelectFromList.show(
        ["{} - {}".format(s.SheetNumber, s.Name) for s in sheet_list],
        title="Select Sheet(s)",
        multiselect=True,
        button_name="Select"
    )

    if not picked:
        return []

    picked_set = set(picked)
    return [
        s for s in sheet_list
        if "{} - {}".format(s.SheetNumber, s.Name) in picked_set
    ]


def check_params_exist(sheet):
    """Warn if any required parameters are missing."""
    missing = []
    for name in [PARAM_SHOW, PARAM_LEFT_DATES, PARAM_LEFT_DESCS,
                 PARAM_RIGHT_DATES, PARAM_RIGHT_DESCS]:
        if get_param(sheet, name) is None:
            missing.append(name)
    return missing


def read_sheet_values(sheet):
    """Read current parameter values from a sheet into a dict."""
    vals = {}
    for name in [PARAM_SHOW, PARAM_LEFT_DATES, PARAM_LEFT_DESCS,
                 PARAM_RIGHT_DATES, PARAM_RIGHT_DESCS]:
        p = get_param(sheet, name)
        if p is None:
            vals[name] = None
        elif p.StorageType == DB.StorageType.Integer:   # Yes/No
            vals[name] = bool(p.AsInteger())
        else:
            vals[name] = p.AsString() or ""
    return vals


def write_sheet_values(sheet, vals, t):
    """Write dict values back to sheet parameters inside an open transaction."""
    for name, value in vals.items():
        p = get_param(sheet, name)
        if p is None or p.IsReadOnly:
            continue
        if p.StorageType == DB.StorageType.Integer:
            p.Set(1 if value else 0)
        else:
            p.Set(value if value is not None else "")


# ---------------------------------------------------------------------------
# WPF Dialog
# ---------------------------------------------------------------------------

XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="WWP Manual Revisions"
    Width="620" Height="580"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    FontFamily="Segoe UI" FontSize="12"
    Background="#F4F4F4">

    <Window.Resources>
        <Style TargetType="Label">
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#333"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="BorderBrush" Value="#BDBDBD"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="AcceptsReturn" Value="True"/>
            <Setter Property="TextWrapping" Value="Wrap"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="16,7"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <StackPanel Grid.Row="0" Margin="0,0,0,14">
            <TextBlock Text="Manual Revision Block" FontSize="16"
                       FontWeight="Bold" Foreground="#1A1A1A"/>
            <TextBlock Text="Editing: " FontSize="11" Foreground="#666"
                       Name="SheetLabel" Margin="0,4,0,0"/>
        </StackPanel>

        <!-- Toggle -->
        <Border Grid.Row="1" Background="#E8F0FE" CornerRadius="6"
                Padding="12,10" Margin="0,0,0,16">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                <CheckBox Name="ChkShowManual" VerticalAlignment="Center"
                          Margin="0,0,10,0"/>
                <StackPanel>
                    <TextBlock Text="Use Manual Revision Block"
                               FontWeight="SemiBold" Foreground="#1A1A1A"/>
                    <TextBlock FontSize="10" Foreground="#555"
                               Text="When checked: manual labels visible, native schedule hidden."/>
                </StackPanel>
            </StackPanel>
        </Border>

        <!-- Two-column input -->
        <Grid Grid.Row="2" Name="InputPanel">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <!-- Left column -->
            <Label Grid.Column="0" Grid.Row="0" Content="LEFT — Dates"/>
            <TextBox Grid.Column="0" Grid.Row="1" Name="TxtLeftDates"
                     MinHeight="80"/>

            <Label Grid.Column="0" Grid.Row="2" Content="LEFT — Descriptions"
                   Margin="0,10,0,0"/>
            <TextBox Grid.Column="0" Grid.Row="3" Name="TxtLeftDescs"
                     MinHeight="80"/>

            <!-- Right column -->
            <Label Grid.Column="2" Grid.Row="0" Content="RIGHT — Dates"/>
            <TextBox Grid.Column="2" Grid.Row="1" Name="TxtRightDates"
                     MinHeight="80"/>

            <Label Grid.Column="2" Grid.Row="2" Content="RIGHT — Descriptions"
                   Margin="0,10,0,0"/>
            <TextBox Grid.Column="2" Grid.Row="3" Name="TxtRightDescs"
                     MinHeight="80"/>
        </Grid>

        <!-- Buttons -->
        <StackPanel Grid.Row="3" Orientation="Horizontal"
                    HorizontalAlignment="Right" Margin="0,16,0,0">
            <TextBlock Name="TxtHint" Foreground="#888" FontSize="10"
                       VerticalAlignment="Center" Margin="0,0,16,0"
                       Text="Tip: one entry per line. Dates and Descs must match line count."/>
            <Button Name="BtnCancel" Content="Cancel"
                    Background="#E0E0E0" Foreground="#333"
                    Margin="0,0,8,0"/>
            <Button Name="BtnApply" Content="Apply to Sheet"
                    Background="#1565C0" Foreground="White"/>
        </StackPanel>

    </Grid>
</Window>
"""


class ManualRevisionDialog(Window):
    def __init__(self, sheet, vals):
        self.sheet = sheet
        self.result_vals = None

        reader = StringReader(XAML)
        win = XamlReader.Load(reader)

        # Wire up named elements
        self.Content      = win.Content
        self.Title        = win.Title
        self.Width        = win.Width
        self.Height       = win.Height
        self.WindowStartupLocation = win.WindowStartupLocation
        self.ResizeMode   = win.ResizeMode
        self.Background   = win.Background
        self.FontFamily   = win.FontFamily
        self.FontSize     = win.FontSize
        self.Resources    = win.Resources

        # Find controls
        def fc(name):
            return win.FindName(name)

        self._sheet_label   = fc("SheetLabel")
        self._chk_show      = fc("ChkShowManual")
        self._txt_l_dates   = fc("TxtLeftDates")
        self._txt_l_descs   = fc("TxtLeftDescs")
        self._txt_r_dates   = fc("TxtRightDates")
        self._txt_r_descs   = fc("TxtRightDescs")
        self._btn_apply     = fc("BtnApply")
        self._btn_cancel    = fc("BtnCancel")

        # Populate
        self._sheet_label.Text = "Editing: {} — {}".format(
            sheet.SheetNumber, sheet.Name)

        self._chk_show.IsChecked    = vals.get(PARAM_SHOW, False)
        self._txt_l_dates.Text      = vals.get(PARAM_LEFT_DATES, "") or ""
        self._txt_l_descs.Text      = vals.get(PARAM_LEFT_DESCS, "") or ""
        self._txt_r_dates.Text      = vals.get(PARAM_RIGHT_DATES, "") or ""
        self._txt_r_descs.Text      = vals.get(PARAM_RIGHT_DESCS, "") or ""

        # Events
        self._btn_apply.Click  += self._on_apply
        self._btn_cancel.Click += self._on_cancel

    def _validate(self):
        """Warn if left date/desc line counts differ."""
        l_dates = self._txt_l_dates.Text.splitlines()
        l_descs = self._txt_l_descs.Text.splitlines()
        r_dates = self._txt_r_dates.Text.splitlines()
        r_descs = self._txt_r_descs.Text.splitlines()

        warnings = []
        if len(l_dates) != len(l_descs):
            warnings.append(
                "LEFT: {} date lines vs {} description lines.".format(
                    len(l_dates), len(l_descs)))
        if len(r_dates) != len(r_descs):
            warnings.append(
                "RIGHT: {} date lines vs {} description lines.".format(
                    len(r_dates), len(r_descs)))

        if warnings:
            msg = "Line count mismatch — rows may not align:\n\n" + \
                  "\n".join(warnings) + "\n\nApply anyway?"
            res = MessageBox.Show(msg, "Line Count Warning",
                                  MessageBoxButton.YesNo)
            return res == MessageBoxResult.Yes
        return True

    def _on_apply(self, sender, e):
        if not self._validate():
            return
        self.result_vals = {
            PARAM_SHOW:        bool(self._chk_show.IsChecked),
            PARAM_LEFT_DATES:  self._txt_l_dates.Text,
            PARAM_LEFT_DESCS:  self._txt_l_descs.Text,
            PARAM_RIGHT_DATES: self._txt_r_dates.Text,
            PARAM_RIGHT_DESCS: self._txt_r_descs.Text,
        }
        self.DialogResult = True
        self.Close()

    def _on_cancel(self, sender, e):
        self.DialogResult = False
        self.Close()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    sheets = get_target_sheets()
    if not sheets:
        forms.alert("No sheets selected or active.", exitscript=True)

    # For multi-sheet selection, warn and confirm
    if len(sheets) > 1:
        res = MessageBox.Show(
            "{} sheets selected.\n\nThe same values will be written to ALL of them.\n\nContinue?".format(len(sheets)),
            "Multiple Sheets",
            MessageBoxButton.YesNo
        )
        if res != MessageBoxResult.Yes:
            return

    # Check params exist on first sheet only (assume same template)
    missing = check_params_exist(sheets[0])
    if missing:
        forms.alert(
            "The following shared parameters are missing from the sheet:\n\n"
            + "\n".join(missing)
            + "\n\nAdd them via Shared Parameters before running this tool.",
            title="Missing Parameters",
            exitscript=True
        )

    # Read values from first sheet to pre-populate dialog
    vals = read_sheet_values(sheets[0])

    # Show dialog
    dlg = ManualRevisionDialog(sheets[0], vals)
    ok = dlg.ShowDialog()

    if not ok or dlg.result_vals is None:
        return

    new_vals = dlg.result_vals

    # Write to all target sheets
    with revit.Transaction("WWP: Manual Revisions"):
        for sheet in sheets:
            write_sheet_values(sheet, new_vals, None)

    forms.toast(
        "Applied to {} sheet{}.".format(
            len(sheets), "s" if len(sheets) > 1 else ""),
        title="Done",
        appid="WWP Revisions"
    )


main()
