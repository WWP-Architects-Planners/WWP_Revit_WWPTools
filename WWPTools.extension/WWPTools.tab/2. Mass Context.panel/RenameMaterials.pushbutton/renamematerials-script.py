#! python3
import clr
import traceback

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, Material, Transaction
from Autodesk.Revit import UI
import WWP_uiUtils as ui


def _get_doc():
    try:
        uidoc = __revit__.ActiveUIDocument
        if uidoc is None:
            return None
        return uidoc.Document
    except Exception:
        return None


def _prompt_find_replace():
    result = ui.uiUtils_find_replace(
        title="Rename Materials",
        find_label="Find:",
        replace_label="Replace:",
        ok_text="OK",
        cancel_text="Cancel",
        width=420,
        height=220,
    )
    if not result:
        return None, None

    find_text = (result.get("find") or "").strip()
    replace_text = result.get("replace") or ""
    if not find_text:
        return None, None

    return find_text, replace_text


def _show_report(title, header_lines, body_lines):
    lines = []
    if header_lines:
        lines.extend(header_lines)
        lines.append("")
    if body_lines:
        lines.extend(body_lines)
    ui.uiUtils_show_text_report(title, "\n".join(lines), ok_text="Close", cancel_text=None, width=720, height=520)


def _show_preview_and_run(doc, find_text, replace_text, planned, skipped):
    lines = []
    lines.append("Found materials to rename:")
    lines.append("Count: {}".format(len(planned)))
    lines.append("Find: {}".format(find_text))
    lines.append("Replace: {}".format(replace_text))
    lines.append("Skipped (conflicts/invalid): {}".format(len(skipped)))
    lines.append("")
    for _, old_name, new_name in planned:
        lines.append("{} -> {}".format(old_name, new_name))
    if skipped:
        lines.append("")
        lines.append("---")
        lines.append("Skipped (preview):")
        for old_name, new_name, reason in skipped[:200]:
            lines.append("{} -> {} [{}]".format(old_name, new_name, reason))

    proceed = ui.uiUtils_show_text_report(
        "Rename Materials - Preview",
        "\n".join(lines),
        ok_text="Proceed",
        cancel_text="Cancel",
        width=760,
        height=560,
    )
    if not proceed:
        return

    renamed, failed = _apply_renames(doc, planned)

    result_lines = []
    result_lines.append("RESULTS")
    result_lines.append("Renamed: {}".format(len(renamed)))
    result_lines.append("Failed: {}".format(len(failed)))
    result_lines.append("Skipped (preview): {}".format(len(skipped)))
    result_lines.append("")

    if renamed:
        result_lines.append("Renamed materials:")
        for old_name, new_name in renamed:
            result_lines.append("{} -> {}".format(old_name, new_name))
    else:
        result_lines.append("Renamed materials:")
        result_lines.append("(none)")

    if failed:
        result_lines.append("")
        result_lines.append("---")
        result_lines.append("Failed (first 20):")
        for old_name, new_name, exc in failed[:20]:
            result_lines.append("{} -> {} ({})".format(old_name, new_name, _format_exception(exc)))

    if skipped:
        result_lines.append("")
        result_lines.append("---")
        result_lines.append("Skipped (first 200):")
        for old_name, new_name, reason in skipped[:200]:
            result_lines.append("{} -> {} [{}]".format(old_name, new_name, reason))

    ui.uiUtils_show_text_report("Rename Materials - Results", "\n".join(result_lines), ok_text="Close", cancel_text=None, width=760, height=560)


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
