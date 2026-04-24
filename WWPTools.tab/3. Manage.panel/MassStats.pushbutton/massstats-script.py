#! python3
"""Mass Stats Tool – live back-of-house calculations from Revit mass elements."""

import clr
import os
import sys
import traceback

from pyrevit import revit

# ── locate the compiled WPF DLL ─────────────────────────────────────────────
script_dir = os.path.dirname(__file__)
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))

def _load_dll():
    """Try net48 first (Revit 2024), fall back to net8.0-windows (Revit 2026)."""
    candidates = [
        os.path.join(lib_path, "WWPTools.WpfUI.net48.dll"),
        os.path.join(lib_path, "WWPTools.WpfUI.net8.0-windows.dll"),
        os.path.join(lib_path, "WWPTools.WpfUI.dll"),
    ]
    for path in candidates:
        if os.path.isfile(path):
            clr.AddReference(path)
            return path
    raise FileNotFoundError(
        "WWPTools.WpfUI DLL not found in {}. "
        "Build the solution first (Build → Build Solution).".format(lib_path)
    )

def main():
    try:
        _load_dll()
    except FileNotFoundError as e:
        from pyrevit import forms
        forms.alert(str(e), title="Mass Stats Tool")
        return

    from WWPTools.WpfUI import MassStatsLauncher  # type: ignore

    try:
        MassStatsLauncher.Show(revit.uiapp)
    except Exception:
        from pyrevit import forms
        forms.alert(traceback.format_exc(), title="Mass Stats Tool – Error")


if __name__ == "__main__":
    main()
