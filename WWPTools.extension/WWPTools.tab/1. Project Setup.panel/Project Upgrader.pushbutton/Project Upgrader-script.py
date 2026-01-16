#! python3
import clr
import os
import traceback

from Autodesk.Revit import DB
import WWP_uiUtils as ui

SILENT_MODE = True

try:
    clr.AddReference('RevitAPI')
    clr.AddReference('RevitAPIUI')
    clr.AddReference('RevitServices')
    import System
    from System.Windows.Forms import (
        Button,
        CheckBox,
        DialogResult,
        Form,
        FormBorderStyle,
        FormStartPosition,
        FolderBrowserDialog,
        Label,
        TextBox,
    )
except Exception:
    Button = None


app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
active_doc = uidoc.Document if uidoc else None


class ProjectUpgraderForm(Form):
    def __init__(self):
        Form.__init__(self)
        self.Text = "Project Upgrader"
        self.Width = 520
        self.Height = 210
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        label = Label()
        label.Text = "Folder"
        label.Location = System.Drawing.Point(12, 20)
        label.Width = 60

        self._folder_input = TextBox()
        self._folder_input.Location = System.Drawing.Point(80, 16)
        self._folder_input.Width = 340

        browse = Button()
        browse.Text = "Browse"
        browse.Location = System.Drawing.Point(430, 14)
        browse.Click += self._browse

        self._subfolders_cb = CheckBox()
        self._subfolders_cb.Text = "Include subfolders"
        self._subfolders_cb.Location = System.Drawing.Point(80, 50)
        self._subfolders_cb.Width = 200
        self._subfolders_cb.Checked = True

        ok = Button()
        ok.Text = "OK"
        ok.Location = System.Drawing.Point(320, 120)
        ok.DialogResult = DialogResult.OK

        cancel = Button()
        cancel.Text = "Cancel"
        cancel.Location = System.Drawing.Point(400, 120)
        cancel.DialogResult = DialogResult.Cancel

        self.AcceptButton = ok
        self.CancelButton = cancel

        self.Controls.Add(label)
        self.Controls.Add(self._folder_input)
        self.Controls.Add(browse)
        self.Controls.Add(self._subfolders_cb)
        self.Controls.Add(ok)
        self.Controls.Add(cancel)

    def _browse(self, sender, args):
        dialog = FolderBrowserDialog()
        dialog.Description = "Select folder containing Revit files"
        if dialog.ShowDialog() == DialogResult.OK:
            self._folder_input.Text = dialog.SelectedPath

    @property
    def folder(self):
        return self._folder_input.Text.strip()

    @property
    def include_subfolders(self):
        return bool(self._subfolders_cb.Checked)


def _pick_options():
    if Button is None:
        if not SILENT_MODE:
            ui.uiUtils_alert("Windows Forms is unavailable on this system.", title="Project Upgrader")
        return None

    form = ProjectUpgraderForm()
    if form.ShowDialog() != DialogResult.OK:
        return None

    return {
        "folder": form.folder,
        "include_subfolders": form.include_subfolders,
    }


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

