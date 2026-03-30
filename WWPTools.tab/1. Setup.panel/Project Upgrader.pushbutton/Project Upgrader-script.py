import clr
import os
import traceback

from Autodesk.Revit import DB
from Autodesk.Revit.DB import FailureProcessingResult
import WWP_uiUtils as ui

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')


app    = __revit__.Application
uidoc  = __revit__.ActiveUIDocument
active_doc = uidoc.Document if uidoc else None


# ─────────────────────────────────────────────────────────────────────────────
# Dialog / failure suppressors (active only during batch processing)
# ─────────────────────────────────────────────────────────────────────────────

def _on_dialog_showing(sender, args):
    """Auto-dismiss any Windows-native dialog boxes."""
    try:
        args.OverrideResult(1)
    except Exception:
        pass


def _on_task_dialog_showing(sender, args):
    """Auto-dismiss Revit TaskDialogs (result 2 = OK)."""
    try:
        args.OverrideResult(2)
    except Exception:
        pass


def _on_failures_processing(sender, args):
    """Delete all warnings and continue so no failure dialog appears."""
    try:
        fa = args.GetFailuresAccessor()
        fa.DeleteAllWarnings()
        args.SetProcessingResult(FailureProcessingResult.Continue)
    except Exception:
        pass


def _suppress_dialogs():
    try:
        __revit__.DialogBoxShowing   += _on_dialog_showing
        __revit__.TaskDialogShowing  += _on_task_dialog_showing
        app.FailuresProcessing       += _on_failures_processing
    except Exception:
        pass


def _restore_dialogs():
    try:
        __revit__.DialogBoxShowing   -= _on_dialog_showing
        __revit__.TaskDialogShowing  -= _on_task_dialog_showing
        app.FailuresProcessing       -= _on_failures_processing
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


def _unload_all_links(doc):
    """Unload all RVT links so they are not reloaded or re-processed during save."""
    try:
        links = list(DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkType))
        if not links:
            return
        t = DB.Transaction(doc, "Unload Links")
        t.Start()
        for lt in links:
            try:
                lt.Unload(None)
            except Exception:
                pass
        t.Commit()
    except Exception:
        pass


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

        _unload_all_links(doc_on_disk)
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
        _unload_all_links(doc_on_disk)
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

    upgraded = []
    skipped  = []
    failed   = []

    _suppress_dialogs()
    try:
        for file_path in files:
            if active_path and os.path.normcase(file_path) == active_path:
                skipped.append(file_path)
                continue

            ext = os.path.splitext(file_path)[1].lower()
            if ext == ".rvt":
                save_path, error = _upgrade_project(file_path, version_suffix)
            else:
                save_path, error = _upgrade_family(file_path, version_suffix)

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
