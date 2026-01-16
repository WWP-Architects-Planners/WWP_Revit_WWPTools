#! python3
import clr
import math
import random
import traceback

from pyrevit import revit, DB
from Autodesk.Revit import UI


PARAM_NAME = "DIM_PlantDiameter"
DEFAULT_PERCENT = 30
PERCENT_MIN = 0
PERCENT_MAX = 100
PERCENT_STEP = 5


try:
    import System
    from System.Drawing import Point, Size
    from System.Windows.Forms import (
        Button,
        CheckBox,
        DialogResult,
        Form,
        FormBorderStyle,
        Label,
        NumericUpDown,
        FormStartPosition,
    )
except Exception:
    System = None


def _get_selected_elements():
    try:
        selection = revit.get_selection()
        return list(selection.elements)
    except Exception:
        try:
            uidoc = __revit__.ActiveUIDocument
            if uidoc is None:
                return []
            return [uidoc.Document.GetElement(eid) for eid in uidoc.Selection.GetElementIds()]
        except Exception:
            return []


def _get_param_value(param):
    if param is None:
        return None
    stype = param.StorageType
    if stype == DB.StorageType.String:
        return param.AsString()
    if stype == DB.StorageType.Double:
        return param.AsDouble()
    if stype == DB.StorageType.Integer:
        return param.AsInteger()
    if stype == DB.StorageType.ElementId:
        return param.AsElementId()
    return None


def _set_param_value(param, value):
    if param is None or param.IsReadOnly:
        return False
    if value is None:
        return False

    stype = param.StorageType
    try:
        if stype == DB.StorageType.String:
            param.Set(str(value))
            return True
        if stype == DB.StorageType.Double:
            param.Set(float(value))
            return True
        if stype == DB.StorageType.Integer:
            param.Set(int(round(value)))
            return True
    except Exception:
        return False

    return False


def _to_number(value):
    if value is None:
        return None
    if isinstance(value, DB.ElementId):
        return None
    try:
        return float(value)
    except Exception:
        return None


def _round_to_nearest_5(number):
    return round(number / 5.0) * 5.0


def _randomize_size(number, percent):
    if number is None:
        return None
    offset = random.uniform(0, percent * number)
    return _round_to_nearest_5(number + offset)


def _set_rotation(doc, element, degrees):
    try:
        loc = element.Location
        if not isinstance(loc, DB.LocationPoint):
            return False
        current = loc.Rotation
        target = math.radians(degrees)
        delta = target - current
        if abs(delta) < 1e-9:
            return True
        axis = DB.Line.CreateUnbound(loc.Point, DB.XYZ.BasisZ)
        DB.ElementTransformUtils.RotateElement(doc, element.Id, axis, delta)
        return True
    except Exception:
        return False


class RandomTreesForm(Form):
    def __init__(self):
        Form.__init__(self)
        self.Text = "Random Tree"
        self.Width = 320
        self.Height = 210
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        self._rotation_cb = CheckBox()
        self._rotation_cb.Text = "Random Rotation"
        self._rotation_cb.Checked = True
        self._rotation_cb.Location = Point(15, 15)
        self._rotation_cb.Width = 260

        self._size_cb = CheckBox()
        self._size_cb.Text = "Random Size"
        self._size_cb.Checked = True
        self._size_cb.Location = Point(15, 45)
        self._size_cb.Width = 260

        self._percent_label = Label()
        self._percent_label.Text = "Size variance (%)"
        self._percent_label.Location = Point(15, 80)
        self._percent_label.Width = 140

        self._percent_input = NumericUpDown()
        self._percent_input.Minimum = System.Decimal(PERCENT_MIN)
        self._percent_input.Maximum = System.Decimal(PERCENT_MAX)
        self._percent_input.Increment = System.Decimal(PERCENT_STEP)
        self._percent_input.Value = System.Decimal(DEFAULT_PERCENT)
        self._percent_input.Location = Point(165, 78)
        self._percent_input.Width = 120

        ok = Button()
        ok.Text = "OK"
        ok.Location = Point(130, 120)
        ok.DialogResult = DialogResult.OK

        cancel = Button()
        cancel.Text = "Cancel"
        cancel.Location = Point(210, 120)
        cancel.DialogResult = DialogResult.Cancel

        self.AcceptButton = ok
        self.CancelButton = cancel

        self.Controls.Add(self._rotation_cb)
        self.Controls.Add(self._size_cb)
        self.Controls.Add(self._percent_label)
        self.Controls.Add(self._percent_input)
        self.Controls.Add(ok)
        self.Controls.Add(cancel)

    @property
    def random_rotation(self):
        return bool(self._rotation_cb.Checked)

    @property
    def random_size(self):
        return bool(self._size_cb.Checked)

    @property
    def percent(self):
        return float(System.Decimal.ToDouble(self._percent_input.Value))


def _ask_for_settings():
    if System is None:
        UI.TaskDialog.Show("Random Tree", "Unable to load input dialog.")
        return None

    form = RandomTreesForm()
    if form.ShowDialog() != DialogResult.OK:
        return None

    return {
        "random_rotation": form.random_rotation,
        "random_size": form.random_size,
        "percent": form.percent,
    }


def main():
    doc = revit.doc
    if doc is None:
        UI.TaskDialog.Show("Random Tree", "No active Revit document found.")
        return

    elements = _get_selected_elements()
    if not elements:
        UI.TaskDialog.Show("Random Tree", "No elements selected. Please select elements and try again.")
        return

    settings = _ask_for_settings()
    if not settings:
        return

    percent = settings["percent"] / 100.0
    random_rotation = settings["random_rotation"]
    random_size = settings["random_size"]

    rotation_targets = list(elements)

    size_elements = []
    size_values = []
    for element in elements:
        param = element.LookupParameter(PARAM_NAME)
        if param is None or param.IsReadOnly:
            continue
        value = _get_param_value(param)
        if isinstance(value, str) and value == "":
            continue
        size_elements.append((element, param))
        size_values.append(value)

    rotation_updates = 0
    rotation_skipped = 0
    size_updates = 0
    size_skipped = 0

    t = DB.Transaction(doc, "Random Tree")
    started = False
    try:
        t.Start()
        started = True

        if random_rotation:
            degrees_list = [random.randint(0, 359 // 5) * 5 for _ in rotation_targets]
        else:
            degrees_list = [0 for _ in rotation_targets]

        for element, degrees in zip(rotation_targets, degrees_list):
            if _set_rotation(doc, element, degrees):
                rotation_updates += 1
            else:
                rotation_skipped += 1

        if random_size:
            for (element, param), value in zip(size_elements, size_values):
                number = _to_number(value)
                new_value = _randomize_size(number, percent) if random_size else number
                if _set_param_value(param, new_value):
                    size_updates += 1
                else:
                    size_skipped += 1
        else:
            size_skipped = len(size_elements)

        t.Commit()
    except Exception as exc:
        if started:
            try:
                t.RollBack()
            except Exception:
                pass
        UI.TaskDialog.Show("Random Tree - Error", "{}\n\n{}".format(exc, traceback.format_exc()))
        return

    summary = [
        "Rotation updated: {}".format(rotation_updates),
        "Rotation skipped: {}".format(rotation_skipped),
        "Size updated: {}".format(size_updates),
        "Size skipped: {}".format(size_skipped),
    ]
    UI.TaskDialog.Show("Random Tree", "\n".join(summary))


if __name__ == "__main__":
    main()
