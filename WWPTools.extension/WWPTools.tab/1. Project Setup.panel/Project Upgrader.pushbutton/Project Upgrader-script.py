#! python3
import clr
import os
import traceback

from Autodesk.Revit import DB
import WWP_uiUtils as ui

SILENT_MODE = True

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')


app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
active_doc = uidoc.Document if uidoc else None


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
    except Exception as exc:
        if not SILENT_MODE:
            ui.uiUtils_alert(str(exc), title="Project Upgrader")
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

    result = []
    for path in files:
        if not os.path.isfile(path):
            continue
        ext = os.path.splitext(path)[1].lower()
        if ext in (".rvt", ".rfa"):
            result.append(path)
    return result


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
    if detach_option is not None:
        options.DetachFromCentralOption = detach_option
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


def main():
    options = _pick_options()
    if not options:
        return

    folder = options["folder"]
    include_subfolders = options["include_subfolders"]
    if not folder or not os.path.isdir(folder):
        if not SILENT_MODE:
            ui.uiUtils_alert("Please select a valid folder.", title="Project Upgrader")
        return

    version_suffix = str(app.VersionNumber)[-2:]
    files = _collect_files(folder, include_subfolders)
    if not files:
        if not SILENT_MODE:
            ui.uiUtils_alert("No Revit files found in the selected folder.", title="Project Upgrader")
        return

    active_path = None
    if active_doc is not None:
        try:
            active_path = os.path.normcase(active_doc.PathName)
        except Exception:
            active_path = None

    upgraded = []
    skipped = []
    failed = []

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

    summary = [
        "Upgraded: {}".format(len(upgraded)),
        "Skipped: {}".format(len(skipped)),
        "Failed: {}".format(len(failed)),
    ]
    if failed:
        summary.append("\nFailures:")
        for path, error in failed[:10]:
            summary.append("- {}: {}".format(path, error))
        if len(failed) > 10:
            summary.append("... {} more".format(len(failed) - 10))

    if not SILENT_MODE:
        ui.uiUtils_alert("\n".join(summary), title="Project Upgrader")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        if not SILENT_MODE:
            ui.uiUtils_alert(traceback.format_exc(), title="Project Upgrader - Error")

