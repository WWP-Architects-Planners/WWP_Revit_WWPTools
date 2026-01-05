#! python3
import clr
import traceback

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, Material, Transaction
from Autodesk.Revit import UI

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (
    Form,
    Label,
    TextBox,
    Button,
    Application,
    ScrollBars,
    DialogResult,
    FormStartPosition,
    FormBorderStyle,
)
from System.Drawing import Point, Size


def _get_doc():
    try:
        uidoc = __revit__.ActiveUIDocument
        if uidoc is None:
            return None
        return uidoc.Document
    except Exception:
        return None


def _prompt_find_replace():
    class _FindReplaceForm(Form):
        def __init__(self):
            # Important under CPython/pythonnet: initialize base Form
            Form.__init__(self)

            self.Text = "Rename Materials"
            self.StartPosition = FormStartPosition.CenterScreen
            self.ClientSize = Size(420, 135)
            self.MinimizeBox = False
            self.MaximizeBox = False
            self.FormBorderStyle = FormBorderStyle.FixedDialog

            lbl_find = Label()
            lbl_find.Text = "Find:"
            lbl_find.Location = Point(12, 15)
            lbl_find.AutoSize = True

            self.txt_find = TextBox()
            self.txt_find.Location = Point(90, 12)
            self.txt_find.Size = Size(310, 20)

            lbl_replace = Label()
            lbl_replace.Text = "Replace:"
            lbl_replace.Location = Point(12, 45)
            lbl_replace.AutoSize = True

            self.txt_replace = TextBox()
            self.txt_replace.Location = Point(90, 42)
            self.txt_replace.Size = Size(310, 20)

            btn_ok = Button()
            btn_ok.Text = "OK"
            btn_ok.Location = Point(244, 85)
            btn_ok.Size = Size(75, 25)
            btn_ok.DialogResult = DialogResult.OK

            btn_cancel = Button()
            btn_cancel.Text = "Cancel"
            btn_cancel.Location = Point(325, 85)
            btn_cancel.Size = Size(75, 25)
            btn_cancel.DialogResult = DialogResult.Cancel

            self.AcceptButton = btn_ok
            self.CancelButton = btn_cancel

            self.Controls.Add(lbl_find)
            self.Controls.Add(self.txt_find)
            self.Controls.Add(lbl_replace)
            self.Controls.Add(self.txt_replace)
            self.Controls.Add(btn_ok)
            self.Controls.Add(btn_cancel)

    try:
        Application.EnableVisualStyles()
    except Exception:
        pass

    dlg = _FindReplaceForm()
    if dlg.ShowDialog() != DialogResult.OK:
        return None, None

    find_text = (dlg.txt_find.Text or "").strip()
    replace_text = dlg.txt_replace.Text or ""
    if not find_text:
        return None, None

    return find_text, replace_text


def _show_report(title, header_lines, body_lines):
    class _ReportForm(Form):
        def __init__(self):
            Form.__init__(self)

            self.Text = title
            self.StartPosition = FormStartPosition.CenterScreen
            self.ClientSize = Size(700, 500)
            self.MinimizeBox = False
            self.MaximizeBox = False
            self.FormBorderStyle = FormBorderStyle.Sizable

            txt = TextBox()
            txt.Multiline = True
            txt.ReadOnly = True
            txt.ScrollBars = ScrollBars.Vertical
            txt.WordWrap = False
            txt.Location = Point(10, 10)
            txt.Size = Size(680, 440)

            btn_close = Button()
            btn_close.Text = "Close"
            btn_close.Location = Point(615, 460)
            btn_close.Size = Size(75, 25)
            btn_close.DialogResult = DialogResult.OK

            self.AcceptButton = btn_close
            self.CancelButton = btn_close

            lines = []
            if header_lines:
                lines.extend(header_lines)
                lines.append("")
            if body_lines:
                lines.extend(body_lines)
            txt.Text = "\r\n".join(lines)

            self.Controls.Add(txt)
            self.Controls.Add(btn_close)

    try:
        Application.EnableVisualStyles()
    except Exception:
        pass

    _ReportForm().ShowDialog()


def _show_preview_and_run(doc, find_text, replace_text, planned, skipped):
    class _PreviewRunForm(Form):
        def __init__(self):
            Form.__init__(self)

            self._has_run = False
            self._planned = planned
            self._skipped = skipped
            self._find_text = find_text
            self._replace_text = replace_text

            self.Text = "Rename Materials"
            self.StartPosition = FormStartPosition.CenterScreen
            self.ClientSize = Size(750, 540)
            self.MinimizeBox = False
            self.MaximizeBox = True
            self.FormBorderStyle = FormBorderStyle.Sizable

            self.txt = TextBox()
            self.txt.Multiline = True
            self.txt.ReadOnly = True
            self.txt.ScrollBars = ScrollBars.Vertical
            self.txt.WordWrap = False
            self.txt.Location = Point(10, 10)
            self.txt.Size = Size(730, 480)

            self.btn_run = Button()
            self.btn_run.Text = "Proceed"
            self.btn_run.Location = Point(584, 505)
            self.btn_run.Size = Size(75, 25)
            self.btn_run.Click += self._on_run

            self.btn_close = Button()
            self.btn_close.Text = "Cancel"
            self.btn_close.Location = Point(665, 505)
            self.btn_close.Size = Size(75, 25)
            self.btn_close.Click += self._on_close

            self.AcceptButton = self.btn_run
            self.CancelButton = self.btn_close

            self.Controls.Add(self.txt)
            self.Controls.Add(self.btn_run)
            self.Controls.Add(self.btn_close)

            self._render_preview()

        def _set_text(self, lines):
            self.txt.Text = "\r\n".join(lines)

        def _render_preview(self):
            lines = []
            lines.append("Found materials to rename:")
            lines.append("Count: {}".format(len(self._planned)))
            lines.append("Find: {}".format(self._find_text))
            lines.append("Replace: {}".format(self._replace_text))
            lines.append("Skipped (conflicts/invalid): {}".format(len(self._skipped)))
            lines.append("")

            for _, old_name, new_name in self._planned:
                lines.append("{} -> {}".format(old_name, new_name))

            if self._skipped:
                lines.append("")
                lines.append("---")
                lines.append("Skipped (preview):")
                for old_name, new_name, reason in self._skipped[:200]:
                    lines.append("{} -> {} [{}]".format(old_name, new_name, reason))

            self._set_text(lines)

        def _on_close(self, sender, args):
            self.Close()

        def _on_run(self, sender, args):
            if self._has_run:
                self.Close()
                return

            self._has_run = True
            self.btn_run.Enabled = False
            self.btn_close.Text = "Close"

            renamed, failed = _apply_renames(doc, self._planned)

            lines = []
            lines.append("RESULTS")
            lines.append("Renamed: {}".format(len(renamed)))
            lines.append("Failed: {}".format(len(failed)))
            lines.append("Skipped (preview): {}".format(len(self._skipped)))
            lines.append("")

            if renamed:
                lines.append("Renamed materials:")
                for old_name, new_name in renamed:
                    lines.append("{} -> {}".format(old_name, new_name))
            else:
                lines.append("Renamed materials:")
                lines.append("(none)")

            if failed:
                lines.append("")
                lines.append("---")
                lines.append("Failed (first 20):")
                for old_name, new_name, exc in failed[:20]:
                    lines.append("{} -> {} ({})".format(old_name, new_name, _format_exception(exc)))

            if self._skipped:
                lines.append("")
                lines.append("---")
                lines.append("Skipped (first 200):")
                for old_name, new_name, reason in self._skipped[:200]:
                    lines.append("{} -> {} [{}]".format(old_name, new_name, reason))

            self._set_text(lines)

    try:
        Application.EnableVisualStyles()
    except Exception:
        pass

    _PreviewRunForm().ShowDialog()


def _format_exception(exc):
    try:
        if hasattr(exc, 'InnerException') and exc.InnerException is not None:
            return "{} | Inner: {}".format(str(exc), str(exc.InnerException))
    except Exception:
        pass
    return str(exc)


def _plan_renames(doc, find_text, replace_text):
    materials = list(FilteredElementCollector(doc).OfClass(Material))
    existing_lower = {m.Name.lower() for m in materials}

    planned = []
    skipped = []

    for mat in materials:
        old_name = mat.Name
        if find_text not in old_name:
            continue

        new_name = (old_name.replace(find_text, replace_text) or "").strip()
        if new_name == old_name:
            continue

        if not new_name:
            skipped.append((old_name, new_name, "empty name"))
            continue

        if len(new_name) > 255:
            skipped.append((old_name, new_name, "name too long"))
            continue

        # Revit often enforces case-insensitive uniqueness.
        if new_name.lower() in existing_lower and new_name.lower() != old_name.lower():
            skipped.append((old_name, new_name, "name conflict (case-insensitive)"))
            continue

        planned.append((mat, old_name, new_name))
        existing_lower.discard(old_name.lower())
        existing_lower.add(new_name.lower())

    return planned, skipped


def _apply_renames(doc, planned):
    renamed = []
    failed = []

    t = Transaction(doc, "Rename Materials")
    started = False
    try:
        t.Start()
        started = True

        for mat, old_name, new_name in planned:
            try:
                mat.Name = new_name
                renamed.append((old_name, new_name))
            except Exception as exc:
                failed.append((old_name, new_name, exc))

        t.Commit()
        return renamed, failed
    except Exception as exc:
        if started:
            try:
                t.RollBack()
            except Exception:
                pass
        # If commit fails, assume changes were not applied.
        return [], failed + [("<transaction>", "<commit>", exc)]


def main():
    doc = _get_doc()
    if doc is None:
        UI.TaskDialog.Show("Rename Materials", "No active Revit document found.")
        return

    find_text, replace_text = _prompt_find_replace()
    if find_text is None:
        UI.TaskDialog.Show("Rename Materials", "Cancelled or no Find text provided.")
        return

    planned, skipped = _plan_renames(doc, find_text, replace_text)
    if not planned:
        UI.TaskDialog.Show(
            "Rename Materials",
            "No matches found.\n\nSkipped: {}".format(len(skipped))
        )
        return

    _show_preview_and_run(doc, find_text, replace_text, planned, skipped)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        UI.TaskDialog.Show("Rename Materials - Error", "{}\n\n{}".format(exc, traceback.format_exc()))
