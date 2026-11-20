#! python3
# PD Export launcher (pyRevit CPython 3). Glue only: loads the net8 DLLs
# and hands control to WWPTools.Export.PdExportLauncher.
# NOTE: pyrevit.forms is NOT supported under CPython; errors are shown with a
# Revit TaskDialog (MessageBox fallback) and the full traceback is printed.
import os
import shutil
import sys
import tempfile
import traceback

import clr

script_dir = os.path.dirname(os.path.abspath(__file__))
# pushbutton -> pulldown -> panel -> tab -> extension root (4 levels up), then lib
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "..", "lib"))

_DLL_NAMES = ("WWPTools.Export.Core.dll", "WWPTools.Export.dll")
_TITLE = "PD Export"


def _show_error(headline, exc, tb_text, dll_paths):
    """Print the full traceback and show it in a dialog. Never raises."""
    dll_text = "\n".join(dll_paths) if dll_paths else "(not resolved)"
    print("{}\n\n{}\n\nDLL paths:\n{}".format(headline, tb_text, dll_text))
    head = "\n".join(tb_text.strip().splitlines()[-12:])
    message = "{}\n\n{}: {}\n\nDLLs:\n{}\n\nTraceback (tail):\n{}".format(
        headline, type(exc).__name__, exc, dll_text, head)
    try:
        clr.AddReference("RevitAPIUI")
        from Autodesk.Revit.UI import TaskDialog
        TaskDialog.Show(_TITLE, message)
        return
    except Exception:
        print("TaskDialog unavailable:\n" + traceback.format_exc())
    try:
        clr.AddReference("PresentationFramework")
        from System.Windows import MessageBox
        MessageBox.Show(message, _TITLE)
    except Exception:
        print("MessageBox unavailable:\n" + traceback.format_exc())


def _is_remote_path(path):
    try:
        import ctypes
        if path.startswith("\\\\"):
            return True
        drive = os.path.splitdrive(path)[0]
        if not drive:
            return False
        return ctypes.windll.kernel32.GetDriveTypeW(drive + "\\") == 4  # DRIVE_REMOTE
    except Exception:
        return False


def _resolve_dll_paths():
    paths = [os.path.join(lib_path, name) for name in _DLL_NAMES]
    for p in paths:
        if not os.path.isfile(p):
            raise IOError("Missing PD Export DLL: {}".format(p))
    if not any(_is_remote_path(p) for p in paths):
        return paths
    # Network share: copy both DLLs to a local temp folder first.
    temp_dir = os.path.join(tempfile.gettempdir(), "WWPTools.Export")
    os.makedirs(temp_dir, exist_ok=True)
    local = []
    for p in paths:
        dest = os.path.join(temp_dir, os.path.basename(p))
        try:
            if (not os.path.isfile(dest)) or os.path.getmtime(dest) < os.path.getmtime(p):
                shutil.copy2(p, dest)
        except Exception:
            if not os.path.isfile(dest):
                dest = p
        local.append(dest)
    return local


def _load_launcher(dll_paths):
    """Load Core first, then Export, by full path; return PdExportLauncher."""
    from System.Reflection import Assembly
    assemblies = []
    for dll in dll_paths:  # order matters: Core before Export
        assemblies.append(Assembly.LoadFrom(dll))
    clr.AddReference(dll_paths[-1])  # make the namespace importable
    from WWPTools.Export import PdExportLauncher
    return PdExportLauncher


def _revit_version_ok():
    """True unless Revit is older than 2025 (the exporters target net8 / Revit 2025+)."""
    try:
        ver = int(str(__revit__.Application.VersionNumber))
    except Exception:
        return True  # unknown version: let the DLL load report any real problem
    if ver >= 2025:
        return True
    message = "PD Export requires Revit 2025 or later. This is Revit {}.".format(ver)
    print(message)
    try:
        clr.AddReference("RevitAPIUI")
        from Autodesk.Revit.UI import TaskDialog
        TaskDialog.Show(_TITLE, message)
    except Exception:
        pass
    return False


def _main():
    if not _revit_version_ok():
        return
    dll_paths = []
    try:
        dll_paths = _resolve_dll_paths()
        launcher = _load_launcher(dll_paths)
    except Exception as exc:
        _show_error("PD Export could not load its assemblies.", exc,
                    traceback.format_exc(), dll_paths)
        return
    try:
        settings_dir = os.path.join(lib_path, "PDExportSettings")
        launcher.Show(__revit__, settings_dir)
    except Exception as exc:
        _show_error("PD Export failed while opening (PdExportLauncher.Show).", exc,
                    traceback.format_exc(), dll_paths)


if __name__ == "__main__" or "__revit__" in globals():
    _main()
