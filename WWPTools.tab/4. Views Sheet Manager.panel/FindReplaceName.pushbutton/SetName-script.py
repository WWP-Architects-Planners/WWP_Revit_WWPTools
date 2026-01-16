#! python3
import importlib
import os
import sys
import traceback

from pyrevit import DB, revit


def load_uiutils():
	script_dir = os.path.dirname(__file__)
	lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))
	if lib_path not in sys.path:
		sys.path.append(lib_path)
	import WWP_uiUtils as ui
	if not hasattr(ui, "uiUtils_viewname_replace_inputs"):
		try:
			ui = importlib.reload(ui)
		except Exception:
			pass
	return ui


def _get_selected_views(doc):
	selection = revit.get_selection()
	selected = list(selection.elements) if selection else []
	views = []
	for element in selected:
		if isinstance(element, DB.Viewport):
			element = doc.GetElement(element.ViewId)
		view = element if isinstance(element, DB.View) else None
		if not view or view.IsTemplate:
			continue
		views.append(view)
	return views


def _build_new_name(current, find_text, replace_text, prefix, suffix):
	new_name = current
	if find_text:
		new_name = new_name.replace(find_text, replace_text)
	if prefix:
		new_name = "{}{}".format(prefix, new_name)
	if suffix:
		new_name = "{}{}".format(new_name, suffix)
	return new_name


def main():
	ui = load_uiutils()
	doc = revit.doc

	views = _get_selected_views(doc)
	if not views:
		ui.uiUtils_alert("Select one or more views first.", title="Replace View Name")
		return

	inputs = ui.uiUtils_viewname_replace_inputs(title="Replace View Name")
	if not inputs:
		return

	find_text = inputs.get("find", "")
	replace_text = inputs.get("replace", "")
	prefix = inputs.get("prefix", "")
	suffix = inputs.get("suffix", "")

	if not any([find_text, replace_text, prefix, suffix]):
		ui.uiUtils_alert("Provide at least one value to apply.", title="Replace View Name")
		return

	transaction = DB.Transaction(doc, "Replace View Names")
	transaction.Start()
	renamed = 0
	failed = []
	try:
		for view in views:
			current = view.Name or ""
			new_name = _build_new_name(current, find_text, replace_text, prefix, suffix)
			if new_name == current:
				continue
			try:
				view.Name = new_name
				renamed += 1
			except Exception as ex:
				failed.append("{} ({})".format(current, ex))
	finally:
		if transaction.HasStarted():
			transaction.Commit()

	message = "Renamed {} view(s).".format(renamed)
	if failed:
		message += "\n\nFailed:\n" + "\n".join(failed[:10])
	if failed and len(failed) > 10:
		message += "\n...and {} more.".format(len(failed) - 10)
	ui.uiUtils_alert(message, title="Replace View Name")


if __name__ == "__main__":
	try:
		main()
	except Exception:
		ui = load_uiutils()
		ui.uiUtils_alert(traceback.format_exc(), title="Replace View Name")
