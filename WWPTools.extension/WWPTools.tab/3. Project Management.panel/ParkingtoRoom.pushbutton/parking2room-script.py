#! python3
import os
import sys
import traceback

import clr
from System import String
from System.Collections.Generic import List

from pyrevit import DB, revit


script_dir = os.path.dirname(__file__)
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))
if lib_path not in sys.path:
    sys.path.append(lib_path)


def _load_uiutils():
    try:
        import WWP_uiUtils as ui
        return ui
    except Exception:
        try:
            from pyrevit import forms
            forms.alert(
                "WWP_uiUtils is not available. Restart pyRevit or reinstall WWPTools.",
                title="Parking Count in Room",
            )
        except Exception:
            pass
        raise


def _get_active_view(doc):
    try:
        return doc.ActiveView
    except Exception:
        return None


def _get_rooms_in_view(doc, view):
    rooms = []
    collector = DB.FilteredElementCollector(doc, view.Id).OfCategory(DB.BuiltInCategory.OST_Rooms)
    for room in collector.WhereElementIsNotElementType():
        if room is None:
            continue
        try:
            if room.Area <= 0:
                continue
        except Exception:
            pass
        rooms.append(room)
    rooms.sort(key=lambda r: ((r.Number or ""), (r.Name or "")))
    return rooms


def _get_parkings_in_view(doc, view):
    collector = DB.FilteredElementCollector(doc, view.Id).OfCategory(DB.BuiltInCategory.OST_Parking)
    return [e for e in collector.WhereElementIsNotElementType() if e is not None]


def _get_location_point(elem):
    try:
        loc = elem.Location
    except Exception:
        loc = None
    if isinstance(loc, DB.LocationPoint):
        return loc.Point
    if isinstance(loc, DB.LocationCurve):
        try:
            return loc.Curve.Evaluate(0.5, True)
        except Exception:
            return None
    return None


def _room_contains_point(room, point):
    if room is None or point is None:
        return False
    try:
        return room.IsPointInRoom(point)
    except Exception:
        return False


def _get_room_solid(doc, room):
    try:
        calc = DB.SpatialElementGeometryCalculator(doc)
        result = calc.CalculateSpatialElementGeometry(room)
        solid = result.GetGeometry()
        return solid
    except Exception:
        return None


def _get_element_solids(elem):
    solids = []
    try:
        opts = DB.Options()
        opts.DetailLevel = DB.ViewDetailLevel.Fine
        opts.ComputeReferences = False
        opts.IncludeNonVisibleObjects = False
        geom = elem.get_Geometry(opts)
    except Exception:
        geom = None
    if geom is None:
        return solids
    for obj in geom:
        solid = None
        if isinstance(obj, DB.Solid):
            solid = obj
        elif isinstance(obj, DB.GeometryInstance):
            try:
                inst_geom = obj.GetInstanceGeometry()
            except Exception:
                inst_geom = None
            if inst_geom is not None:
                for inst_obj in inst_geom:
                    if isinstance(inst_obj, DB.Solid):
                        solids.append(inst_obj)
                continue
        if solid is not None:
            solids.append(solid)
    solids = [s for s in solids if s is not None and s.Volume > 1e-9]
    return solids


def _sum_solid_volume(solids):
    total = 0.0
    for solid in solids:
        try:
            total += solid.Volume
        except Exception:
            continue
    return total


def _intersection_volume(a_solid, b_solid):
    try:
        result = DB.BooleanOperationsUtils.ExecuteBooleanOperation(
            a_solid, b_solid, DB.BooleanOperationsType.Intersect
        )
    except Exception:
        return 0.0
    try:
        return result.Volume
    except Exception:
        return 0.0


def _get_room_label(room):
    number = ""
    name = ""
    try:
        number = room.Number or ""
    except Exception:
        number = ""
    try:
        name = room.Name or ""
    except Exception:
        name = ""
    label = ""
    if number and name:
        label = "{} - {}".format(number, name)
    else:
        label = number or name or "Room"
    return "{} (Id:{})".format(label, room.Id.IntegerValue)


def _collect_room_parking_map(rooms, parkings):
    room_map = {r.Id.IntegerValue: [] for r in rooms}
    room_solids = {}
    for room in rooms:
        room_solids[room.Id.IntegerValue] = _get_room_solid(room.Document, room)

    for parking in parkings:
        assigned = False
        solids = _get_element_solids(parking)
        total_volume = _sum_solid_volume(solids)
        if solids and total_volume > 0:
            best_room_id = None
            best_volume = 0.0
            for room in rooms:
                room_solid = room_solids.get(room.Id.IntegerValue)
                if room_solid is None:
                    continue
                intersection = 0.0
                for solid in solids:
                    intersection += _intersection_volume(solid, room_solid)
                if intersection > best_volume:
                    best_volume = intersection
                    best_room_id = room.Id.IntegerValue
            if best_room_id is not None and best_volume > 0:
                room_map[best_room_id].append(parking)
                assigned = True

        if assigned:
            continue

        point = _get_location_point(parking)
        if point is None:
            continue
        for room in rooms:
            if _room_contains_point(room, point):
                room_map[room.Id.IntegerValue].append(parking)
                break
    return room_map


def _get_param_names(elements):
    names = set()
    for elem in elements:
        try:
            for param in elem.Parameters:
                if param is None:
                    continue
                try:
                    pname = param.Definition.Name
                except Exception:
                    pname = None
                if pname:
                    names.add(pname)
        except Exception:
            continue
    return sorted(names, key=lambda n: n.lower())


def _get_type_param_names(elements):
    names = set()
    for elem in elements:
        try:
            elem_type = elem.Document.GetElement(elem.GetTypeId())
        except Exception:
            elem_type = None
        if elem_type is None:
            continue
        try:
            for param in elem_type.Parameters:
                if param is None:
                    continue
                try:
                    pname = param.Definition.Name
                except Exception:
                    pname = None
                if pname:
                    names.add(pname)
        except Exception:
            continue
    return sorted(names, key=lambda n: n.lower())


def _get_string_room_params(room):
    options = []
    for param in room.Parameters:
        if param is None or param.IsReadOnly:
            continue
        try:
            if param.StorageType != DB.StorageType.String:
                continue
        except Exception:
            continue
        try:
            name = param.Definition.Name
        except Exception:
            name = None
        if name:
            options.append(name)
    options = sorted(set(options), key=lambda n: n.lower())
    return options


def _get_param_value(param, doc):
    if param is None:
        return None
    stype = param.StorageType
    try:
        if stype == DB.StorageType.String:
            return param.AsString()
        if stype == DB.StorageType.Integer:
            return param.AsInteger()
        if stype == DB.StorageType.Double:
            return param.AsDouble()
        if stype == DB.StorageType.ElementId:
            elem_id = param.AsElementId()
            if elem_id and elem_id.IntegerValue > 0:
                try:
                    elem = doc.GetElement(elem_id)
                    if elem:
                        return elem.Name
                except Exception:
                    pass
            return elem_id.IntegerValue
    except Exception:
        return None
    return None


def _get_param_by_name(elem, name):
    try:
        return elem.LookupParameter(name)
    except Exception:
        return None


def _get_parking_type_key(parking, doc, mode, param_name):
    if mode == "family_type":
        try:
            symbol = doc.GetElement(parking.GetTypeId())
            if symbol and symbol.Name:
                return symbol.Name
        except Exception:
            return "Type"
        return "Type"
    if mode == "instance_param":
        return _param_as_string(_get_param_by_name(parking, param_name), doc)
    if mode == "type_param":
        try:
            symbol = doc.GetElement(parking.GetTypeId())
        except Exception:
            symbol = None
        return _param_as_string(_get_param_by_name(symbol, param_name), doc)
    return "Type"


def _param_as_string(param, doc):
    value = _get_param_value(param, doc)
    if value is None:
        return ""
    return str(value)


def _get_parking_count(parking, doc, count_param_name=None, count_param_mode="instance"):
    if not count_param_name:
        return 1
    if count_param_mode == "type":
        try:
            symbol = doc.GetElement(parking.GetTypeId())
        except Exception:
            symbol = None
        count_param = _get_param_by_name(symbol, count_param_name)
    else:
        count_param = _get_param_by_name(parking, count_param_name)
    if count_param is None:
        return 1
    value = _get_param_value(count_param, doc)
    if value is None:
        return 1
    try:
        return int(value)
    except Exception:
        return 1


def _format_counts(type_counts):
    if not type_counts:
        return "N/A"
    parts = []
    for key in sorted(type_counts.keys(), key=lambda k: k.lower()):
        parts.append("{} :{}".format(key, type_counts[key]))
    return "\t".join(parts)


def _show_inputs_form(
    rooms,
    type_options,
    room_param_options,
    count_param_options,
    default_room_param,
):
    clr.AddReference("PresentationFramework")
    clr.AddReference("PresentationCore")
    clr.AddReference("WindowsBase")
    from System.IO import StringReader
    from System.Windows.Markup import XamlReader
    from System.Xml import XmlReader
    from System.Windows import Window
    from System.Windows.Controls import SelectionMode

    def _to_net_list(items):
        net_list = List[String]()
        for item in items:
            net_list.Add("" if item is None else str(item))
        return net_list

    xaml = """
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="Parking Count in Room" Height="620" Width="720"
            WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip">
        <Grid Margin="12">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <StackPanel Grid.Row="0">
                <TextBlock Text="Select rooms that have parking:" Margin="0,0,0,6"/>
            </StackPanel>
            <ListBox Name="RoomsList" Grid.Row="1" SelectionMode="Extended" />
            <StackPanel Grid.Row="2" Margin="0,10,0,0">
                <TextBlock Text="Parking type source:" />
                <ComboBox Name="TypeSourceCombo" Margin="0,4,0,8"/>
                <TextBlock Text="Room parameter to write:" />
                <ComboBox Name="RoomParamCombo" Margin="0,4,0,8"/>
                <TextBlock Text="Parking count source parameter:" />
                <ComboBox Name="CountParamCombo" Margin="0,4,0,12"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button Name="OkButton" Content="Run" Width="90" Margin="0,0,8,0"/>
                    <Button Name="CancelButton" Content="Cancel" Width="90"/>
                </StackPanel>
            </StackPanel>
            <Image Name="LogoImage" Width="64" Height="64"
                   HorizontalAlignment="Right" VerticalAlignment="Bottom"
                   Grid.Row="2" Margin="0,0,0,0" />
        </Grid>
    </Window>
    """
    reader = XmlReader.Create(StringReader(xaml))
    window = XamlReader.Load(reader)

    rooms_list = window.FindName("RoomsList")
    type_combo = window.FindName("TypeSourceCombo")
    room_combo = window.FindName("RoomParamCombo")
    count_combo = window.FindName("CountParamCombo")
    ok_button = window.FindName("OkButton")
    cancel_button = window.FindName("CancelButton")
    logo_image = window.FindName("LogoImage")

    rooms_list.ItemsSource = _to_net_list(rooms)
    type_combo.ItemsSource = _to_net_list(type_options)
    room_combo.ItemsSource = _to_net_list(room_param_options)
    count_combo.ItemsSource = _to_net_list(count_param_options)

    if type_options:
        type_combo.SelectedIndex = 0
    if room_param_options:
        try:
            room_combo.SelectedItem = default_room_param
        except Exception:
            room_combo.SelectedIndex = 0
    if count_param_options:
        count_combo.SelectedIndex = 0

    try:
        from System import Uri
        from System.Windows.Media.Imaging import BitmapImage
        logo_path = os.path.join(lib_path, "WWPtools-logo.png")
        if os.path.isfile(logo_path):
            bitmap = BitmapImage()
            bitmap.BeginInit()
            bitmap.UriSource = Uri(logo_path)
            bitmap.CacheOption = BitmapImage.CacheOption.OnLoad
            bitmap.EndInit()
            logo_image.Source = bitmap
    except Exception:
        pass

    def _ok(_sender, _args):
        window.DialogResult = True
        window.Close()

    def _cancel(_sender, _args):
        window.DialogResult = False
        window.Close()

    ok_button.Click += _ok
    cancel_button.Click += _cancel

    if not window.ShowDialog():
        return None

    selected_rooms = [str(item) for item in rooms_list.SelectedItems]
    return {
        "rooms": selected_rooms,
        "type_source": str(type_combo.SelectedItem) if type_combo.SelectedItem is not None else None,
        "room_param": str(room_combo.SelectedItem) if room_combo.SelectedItem is not None else None,
        "count_param": str(count_combo.SelectedItem) if count_combo.SelectedItem is not None else None,
    }


def main():
    ui = _load_uiutils()
    doc = revit.doc
    if doc is None:
        ui.uiUtils_alert("No active document.", title="Parking Count in Room")
        return

    view = _get_active_view(doc)
    if view is None:
        ui.uiUtils_alert("No active view.", title="Parking Count in Room")
        return

    rooms = _get_rooms_in_view(doc, view)
    if not rooms:
        ui.uiUtils_alert("No rooms found in the current view.", title="Parking Count in Room")
        return

    parkings = _get_parkings_in_view(doc, view)
    if not parkings:
        ui.uiUtils_alert("No parking elements found in the current view.", title="Parking Count in Room")
        return

    room_parking = _collect_room_parking_map(rooms, parkings)
    rooms_with_parking = [r for r in rooms if room_parking.get(r.Id.IntegerValue)]
    if not rooms_with_parking:
        ui.uiUtils_alert("No rooms with parking found in the current view.", title="Parking Count in Room")
        return

    instance_param_names = _get_param_names(parkings)
    type_param_names = _get_type_param_names(parkings)
    type_options = ["Family Type"] + ["Type Parameter: " + n for n in type_param_names] + [
        "Instance Parameter: " + n for n in instance_param_names
    ]

    room_param_options = _get_string_room_params(rooms_with_parking[0])
    if not room_param_options:
        ui.uiUtils_alert("No writable string parameters found on rooms.", title="Parking Count in Room")
        return

    default_room_param = "Parking Count"
    count_param_options = ["(Default 1)"] + ["Type Parameter: " + n for n in type_param_names] + [
        "Instance Parameter: " + n for n in instance_param_names
    ]
    room_labels = [_get_room_label(r) for r in rooms_with_parking]
    room_lookup = dict(zip(room_labels, rooms_with_parking))

    default_room_param_value = default_room_param if default_room_param in room_param_options else room_param_options[0]
    inputs = _show_inputs_form(
        room_labels,
        type_options,
        room_param_options,
        count_param_options,
        default_room_param_value,
    )
    if not inputs:
        return

    selected_room_labels = inputs.get("rooms") or []
    selected_rooms = [room_lookup[lbl] for lbl in selected_room_labels if lbl in room_lookup]
    if not selected_rooms:
        return

    selected_type_option = inputs.get("type_source")
    if not selected_type_option:
        return
    if selected_type_option == "Family Type":
        type_mode = "family_type"
        type_param_name = None
    elif selected_type_option.startswith("Type Parameter: "):
        type_mode = "type_param"
        type_param_name = selected_type_option.replace("Type Parameter: ", "", 1)
    else:
        type_mode = "instance_param"
        type_param_name = selected_type_option.replace("Instance Parameter: ", "", 1)

    target_room_param = inputs.get("room_param")
    if not target_room_param:
        return

    selected_count_param = inputs.get("count_param")
    if not selected_count_param:
        return
    if selected_count_param == "(Default 1)":
        count_param_name = None
        count_param_mode = "instance"
    elif selected_count_param.startswith("Type Parameter: "):
        count_param_name = selected_count_param.replace("Type Parameter: ", "", 1)
        count_param_mode = "type"
    else:
        count_param_name = selected_count_param.replace("Instance Parameter: ", "", 1)
        count_param_mode = "instance"

    updated = 0
    skipped = 0
    failures = 0
    tx = DB.Transaction(doc, "Parking Count in Room")
    try:
        tx.Start()
        for room in selected_rooms:
            parking_list = room_parking.get(room.Id.IntegerValue, [])
            if not parking_list:
                skipped += 1
                continue
            type_counts = {}
            for parking in parking_list:
                key = _get_parking_type_key(parking, doc, type_mode, type_param_name) or "Type"
                count = _get_parking_count(
                    parking,
                    doc,
                    count_param_name=count_param_name,
                    count_param_mode=count_param_mode,
                )
                type_counts[key] = type_counts.get(key, 0) + count

            value = _format_counts(type_counts)
            param = _get_param_by_name(room, target_room_param)
            if param is None or param.IsReadOnly:
                failures += 1
                continue
            try:
                param.Set(value)
                updated += 1
            except Exception:
                failures += 1
        tx.Commit()
    except Exception:
        try:
            tx.RollBack()
        except Exception:
            pass
        ui.uiUtils_alert(traceback.format_exc(), title="Parking Count in Room")
        return

    report = [
        "Rooms updated: {}".format(updated),
        "Rooms skipped: {}".format(skipped),
        "Failures: {}".format(failures),
        "",
        "Type source: {}".format(selected_type_option),
        "Room parameter: {}".format(target_room_param),
        "Count parameter: {}".format(selected_count_param),
    ]
    ui.uiUtils_show_text_report(
        "Parking Count in Room - Results",
        "\n".join(report),
        ok_text="Close",
        cancel_text=None,
        width=620,
        height=360,
    )


if __name__ == "__main__":
    try:
        main()
    except Exception:
        ui = _load_uiutils()
        ui.uiUtils_alert(traceback.format_exc(), title="Parking Count in Room")
