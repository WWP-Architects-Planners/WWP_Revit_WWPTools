import clr
import os
import traceback

from Autodesk.Revit import DB
from Autodesk.Revit.DB import FailureProcessingResult
import WWP_uiUtils as ui

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')


app        = __revit__.Application
uidoc      = __revit__.ActiveUIDocument
active_doc = uidoc.Document if uidoc else None


# ─────────────────────────────────────────────────────────────────────────────
# Pre-scan: detect local links by reading raw bytes (no Revit open needed)
# ─────────────────────────────────────────────────────────────────────────────

def _scan_local_links(file_path, candidate_basenames):
    """
    Read the file's binary content and search for UTF-16-LE encoded filenames.
    Revit stores linked file paths as plain UTF-16-LE strings in the OLE binary,
    so a raw byte search is reliable without needing to open the file in Revit.
    Returns the subset of candidate_basenames found in this file.
    """
    found = set()
    try:
        with open(file_path, 'rb') as f:
            data = f.read()
        for name in candidate_basenames:
            if name.encode('utf-16-le') in data:
                found.add(name)
    except Exception:
        pass
    return found


def _build_link_map(files):
    """
    For each file, find which other files in the same batch it links to.
    Returns dict: file_path -> [linked_file_paths]
    Scan is done on raw bytes so no Revit open is required.
    """
    basename_to_path = {}
    for fp in files:
        bn = os.path.basename(fp)
        basename_to_path[bn.lower()] = fp   # key is lowercase for matching

    link_map = {}
    for fp in files:
        me = os.path.basename(fp).lower()
        candidates = [bn for bn in basename_to_path if bn != me]
        found_lower = _scan_local_links(fp, candidates)
        link_map[fp] = [basename_to_path[n] for n in found_lower]

    return link_map


# ─────────────────────────────────────────────────────────────────────────────
# Temporary rename helpers
# ─────────────────────────────────────────────────────────────────────────────

_HIDDEN_EXT = ".upgrade_hidden"


def _hide_links(linked_paths):
    """
    Rename each linked file to <name>.upgrade_hidden so Revit cannot find it.
    Returns dict of {original_path: temp_path} for files successfully renamed.
    """
    hidden = {}
    for lp in linked_paths:
        if not os.path.exists(lp):
            continue
        temp = lp + _HIDDEN_EXT
        try:
            os.rename(lp, temp)
            hidden[lp] = temp
        except Exception:
            pass
    return hidden


def _restore_links(hidden):
    """Rename temp files back to their original names."""
    for orig, temp in hidden.items():
        if os.path.exists(temp):
            try:
                os.rename(temp, orig)
            except Exception:
                pass


# ─────────────────────────────────────────────────────────────────────────────
# Dialog / failure suppressors (active only during batch processing)
# ─────────────────────────────────────────────────────────────────────────────

def _on_dialog_showing(sender, args):
    try:
        args.OverrideResult(1)
    except Exception:
        pass


def _on_task_dialog_showing(sender, args):
    try:
        args.OverrideResult(2)
    except Exception:
        pass


def _on_failures_processing(sender, args):
    try:
        fa = args.GetFailuresAccessor()
        fa.DeleteAllWarnings()
        args.SetProcessingResult(FailureProcessingResult.Continue)
    except Exception:
        pass


def _suppress_dialogs():
    try:
        __revit__.DialogBoxShowing  += _on_dialog_showing
        __revit__.TaskDialogShowing += _on_task_dialog_showing
        app.FailuresProcessing      += _on_failures_processing
    except Exception:
        pass


def _restore_dialogs():
    try:
        __revit__.DialogBoxShowing  -= _on_dialog_showing
        __revit__.TaskDialogShowing -= _on_task_dialog_showing
        app.FailuresProcessing      -= _on_failures_processing
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# Core helpers
# ─────────────────────────────────────────────────────────────────────────────

def _pick_options():
    try:
        return ui.uiUtils_project_upgrader_options(
            title="Project Upgrader",
            description="Select folder containing Revit files",
            include_subfolders_label="Include subfolders",
            cancel_text="Cancel",
            width=520,
            height=260,
            initial_folder="",
        )
    except Exception:
        return None


def _collect_files(folder, include_subfolders):
    files = []
    if include_subfolders:
        for root, dirs, filenames in os.walk(folder):
            for name in filenames:
                files.append(os.path.join(root, name))
    else:
        for name in os.listdir(folder):
            files.append(os.path.join(folder, name))

    return [p for p in files
            if os.path.isfile(p) and os.path.splitext(p)[1].lower() in (".rvt", ".rfa")]


def _build_save_path(file_path, version_suffix):
    base, ext = os.path.splitext(file_path)
    candidate = "{}_R{}{}".format(base, version_suffix, ext)
    if not os.path.exists(candidate):
        return candidate
    index = 2
    while True:
        candidate = "{}_R{}_{}{}".format(base, version_suffix, index, ext)
        if not os.path.exists(candidate):
            return candidate
        index += 1


def _open_document(file_path, detach_option=None):
    model_path = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(file_path)
    options = DB.OpenOptions()
    options.AuditOnOpen = False
    if detach_option is not None:
        options.DetachFromCentralOption = detach_option
    try:
        ws_config = DB.WorksetConfiguration(DB.WorksetConfigurationOption.CloseAllWorksets)
        options.SetOpenWorksetsConfiguration(ws_config)
    except Exception:
        pass
    return app.OpenDocumentFile(model_path, options)


def _upgrade_project(file_path, version_suffix):
    doc_on_disk = None
    try:
        try:
            doc_on_disk = _open_document(file_path, DB.DetachFromCentralOption.DetachAndPreserveWorksets)
        except Exception:
            doc_on_disk = _open_document(file_path, DB.DetachFromCentralOption.DoNotDetach)

        save_path = _build_save_path(file_path, version_suffix)
        save_options = DB.SaveAsOptions()
        save_options.OverwriteExistingFile = False
        try:
            if doc_on_disk.IsWorkshared:
                ws_options = DB.WorksharingSaveAsOptions()
                ws_options.SaveAsCentral = True
                save_options.SetWorksharingOptions(ws_options)
        except Exception:
            pass
        doc_on_disk.SaveAs(save_path, save_options)
        return save_path, None
    except Exception as exc:
        return None, str(exc)
    finally:
        if doc_on_disk is not None:
            try:
                doc_on_disk.Close(True)
            except Exception:
                pass


def _upgrade_family(file_path, version_suffix):
    doc_on_disk = None
    try:
        doc_on_disk = _open_document(file_path, None)
        save_path = _build_save_path(file_path, version_suffix)
        save_options = DB.SaveAsOptions()
        save_options.OverwriteExistingFile = False
        doc_on_disk.SaveAs(save_path, save_options)
        return save_path, None
    except Exception as exc:
        return None, str(exc)
    finally:
        if doc_on_disk is not None:
            try:
                doc_on_disk.Close(True)
            except Exception:
                pass


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    options = _pick_options()
    if not options:
        return

    folder = options["folder"]
    include_subfolders = options["include_subfolders"]
    if not folder or not os.path.isdir(folder):
        ui.uiUtils_alert("Please select a valid folder.", title="Project Upgrader")
        return

    version_suffix = str(app.VersionNumber)[-2:]
    files = _collect_files(folder, include_subfolders)
    if not files:
        ui.uiUtils_alert("No Revit files found in the selected folder.", title="Project Upgrader")
        return

    active_path = None
    if active_doc is not None:
        try:
            active_path = os.path.normcase(active_doc.PathName)
        except Exception:
            pass

    # Scan all files BEFORE opening any of them to find which link which.
    # Files referenced by another file get temporarily renamed so Revit
    # cannot load them as links when the host file is opened.
    link_map = _build_link_map(files)

    upgraded = []
    skipped  = []
    failed   = []

    _suppress_dialogs()
    try:
        for file_path in files:
            if active_path and os.path.normcase(file_path) == active_path:
                skipped.append(file_path)
                continue

            # Hide linked files so Revit skips loading them on open
            hidden = _hide_links(link_map.get(file_path, []))
            try:
                ext = os.path.splitext(file_path)[1].lower()
                if ext == ".rvt":
                    save_path, error = _upgrade_project(file_path, version_suffix)
                else:
                    save_path, error = _upgrade_family(file_path, version_suffix)
            finally:
                # Always restore before moving to the next file
                _restore_links(hidden)

            if save_path:
                upgraded.append(save_path)
            else:
                failed.append((file_path, error or "Unknown error"))
    finally:
        _restore_dialogs()

    summary = [
        "Upgraded: {}".format(len(upgraded)),
        "Skipped:  {}".format(len(skipped)),
        "Failed:   {}".format(len(failed)),
    ]
    if failed:
        summary.append("\nFailures:")
        for path, error in failed[:10]:
            summary.append("  {}\n  -> {}".format(os.path.basename(path), error))
        if len(failed) > 10:
            summary.append("  ... {} more".format(len(failed) - 10))

    ui.uiUtils_alert("\n".join(summary), title="Project Upgrader")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        ui.uiUtils_alert(traceback.format_exc(), title="Project Upgrader - Error")
