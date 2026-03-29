#!python3
"""Redirects to the combined Bulk Rename tool (pre-selects Materials)."""
import os
import importlib.util
import traceback

script_dir = os.path.dirname(__file__)
_target = os.path.abspath(
    os.path.join(
        script_dir, "..", "..",
        "3. Manage.panel", "FindReplaceName.pushbutton", "SetName-script.py"
    )
)

try:
    spec = importlib.util.spec_from_file_location("bulkrename", _target)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod.main()
except Exception:
    from Autodesk.Revit.UI import TaskDialog
    TaskDialog.Show("Bulk Rename – Error", traceback.format_exc())
