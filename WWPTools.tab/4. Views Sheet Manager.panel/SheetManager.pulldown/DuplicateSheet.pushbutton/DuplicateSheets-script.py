#! python3
import importlib
import os
import sys
import traceback

from pyrevit import DB, revit


def load_uiutils():
	script_dir = os.path.dirname(__file__)
	lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "..", "lib"))
	if lib_path not in sys.path:
		sys.path.append(lib_path)
	import WWP_uiUtils as ui
	if not hasattr(ui, "uiUtils_duplicate_sheet_inputs"):
		try:
			ui = importlib.reload(ui)
		except Exception:
			pass
	return ui


def _collect_sheets(doc):
	sheets = []
	for view in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet):
		sheets.append(view)
	return sorted(sheets, key=lambda s: s.SheetNumber or "")


def _get_titleblock_type_id(doc, sheet):
	titleblocks = (
		DB.FilteredElementCollector(doc, sheet.Id)
		.OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
		.WhereElementIsNotElementType()
		.ToElements()
	)
	if titleblocks:
		return titleblocks[0].GetTypeId()
	return DB.ElementId.InvalidElementId


def _collect_viewports(doc, sheet):
	viewports = []
	for viewport in DB.FilteredElementCollector(doc, sheet.Id).OfClass(DB.Viewport):
		view = doc.GetElement(viewport.ViewId)
		if view:
			viewports.append((view, viewport.GetBoxCenter()))
	return viewports


def _collect_schedule_instances(doc, sheet):
	instances = []
	for inst in DB.FilteredElementCollector(doc, sheet.Id).OfClass(DB.ScheduleSheetInstance):
		view = doc.GetElement(inst.ScheduleId)
		if view:
			instances.append((view, inst.Point))
	return instances


def _duplicate_view(view, option):
	return view.Duplicate(option)


def _apply_prefix_suffix(value, prefix, suffix):
	updated = value
	if prefix:
		updated = "{}{}".format(prefix, updated)
	if suffix:
		updated = "{}{}".format(updated, suffix)
	return updated


def _map_duplicate_option(option_index):
	if option_index == 1:
		return DB.ViewDuplicateOption.WithDetailing
	if option_index == 2:
		return DB.ViewDuplicateOption.AsDependent
	return DB.ViewDuplicateOption.Duplicate


def main():
	ui = load_uiutils()
	doc = revit.doc

	sheets = _collect_sheets(doc)
	if not sheets:
		ui.uiUtils_alert("No sheets found.", title="Sheet Duplicator")
		return

	display_items = [
		"{} - {}".format(sheet.SheetNumber or "", sheet.Name or "")
		for sheet in sheets
	]
	inputs = ui.uiUtils_duplicate_sheet_inputs(display_items, title="Sheet Duplicator")
	if not inputs:
		return

	selected_indices = inputs.get("selected_indices") or []
	if not selected_indices:
		return

	duplicate_with_views = inputs.get("duplicate_with_views", True)
	duplicate_option = _map_duplicate_option(inputs.get("duplicate_option", 0))
	prefix = inputs.get("prefix", "")
	suffix = inputs.get("suffix", "")

	selected_sheets = [sheets[i] for i in selected_indices]
	created_sheets = []
	failed = []

	transaction = DB.Transaction(doc, "Duplicate Sheets")
	transaction.Start()
	try:
		for sheet in selected_sheets:
			try:
				titleblock_type_id = _get_titleblock_type_id(doc, sheet)
				new_sheet = DB.ViewSheet.Create(doc, titleblock_type_id)
				new_number = _apply_prefix_suffix(sheet.SheetNumber or "", prefix, suffix)
				if new_number:
					new_sheet.SheetNumber = new_number
				new_sheet.Name = sheet.Name or ""

				if duplicate_with_views:
					viewports = _collect_viewports(doc, sheet)
					for view, point in viewports:
						try:
							new_view_id = _duplicate_view(view, duplicate_option)
							new_view = doc.GetElement(new_view_id)
							if new_view and (prefix or suffix):
								new_view.Name = _apply_prefix_suffix(new_view.Name or "", prefix, suffix)
							DB.Viewport.Create(doc, new_sheet.Id, new_view_id, point)
						except Exception as ex:
							failed.append("{} (viewport: {})".format(sheet.SheetNumber, ex))

					schedules = _collect_schedule_instances(doc, sheet)
					for view, point in schedules:
						try:
							new_view_id = _duplicate_view(view, duplicate_option)
							new_view = doc.GetElement(new_view_id)
							if new_view and (prefix or suffix):
								new_view.Name = _apply_prefix_suffix(new_view.Name or "", prefix, suffix)
							DB.ScheduleSheetInstance.Create(doc, new_sheet.Id, new_view_id, point)
						except Exception as ex:
							failed.append("{} (schedule: {})".format(sheet.SheetNumber, ex))

				created_sheets.append(new_sheet)
			except Exception as ex:
				failed.append("{} ({})".format(sheet.SheetNumber, ex))
	finally:
		if transaction.HasStarted():
			transaction.Commit()

	message = "Created {} sheet(s).".format(len(created_sheets))
	if failed:
		message += "\n\nFailed:\n" + "\n".join(failed[:10])
	if failed and len(failed) > 10:
		message += "\n...and {} more.".format(len(failed) - 10)
	ui.uiUtils_alert(message, title="Sheet Duplicator")


if __name__ == "__main__":
	try:
		main()
	except Exception:
		ui = load_uiutils()
		ui.uiUtils_alert(traceback.format_exc(), title="Sheet Duplicator")
