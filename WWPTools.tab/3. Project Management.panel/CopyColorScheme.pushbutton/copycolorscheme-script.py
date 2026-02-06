# -*- coding: utf-8 -*-

from pyrevit import revit, DB
import WWP_uiUtils as ui

try:
	import clr
	from System.Collections.Generic import List
except Exception:
	clr = None
	List = None


def _collect_color_fill_schemes(doc):
	# ColorFillScheme exists in Revit 2022+
	schemes = list(DB.FilteredElementCollector(doc).OfClass(DB.ColorFillScheme).ToElements())
	# Sort by name for nicer UX
	schemes.sort(key=lambda s: (getattr(s, "Name", "") or "").lower())
	return schemes


def _unique_scheme_name(existing_names, desired_name):
	base = (desired_name or "").strip()
	if not base:
		base = "New Color Scheme"
	name = base
	i = 2
	while name.lower() in existing_names:
		name = "{} ({})".format(base, i)
		i += 1
	return name


def _call_if_exists(obj, method_name, *args):
	method = getattr(obj, method_name, None)
	if callable(method):
		return method(*args)
	return None


def _scheme_area_scheme_name(doc, scheme):
	try:
		area_scheme_id = getattr(scheme, "AreaSchemeId", None)
		if area_scheme_id:
			area_scheme = doc.GetElement(area_scheme_id)
			if area_scheme:
				return getattr(area_scheme, "Name", "") or ""
	except Exception:
		pass
	try:
		get_area_scheme_id = getattr(scheme, "GetAreaSchemeId", None)
		if callable(get_area_scheme_id):
			area_scheme_id = get_area_scheme_id()
			if area_scheme_id:
				area_scheme = doc.GetElement(area_scheme_id)
				if area_scheme:
					return getattr(area_scheme, "Name", "") or ""
	except Exception:
		pass
	try:
		param = scheme.LookupParameter("Area Scheme")
		if param:
			return param.AsString() or param.AsValueString() or ""
	except Exception:
		pass
	return ""


def _scheme_display_name(doc, scheme):
	scheme_name = getattr(scheme, "Name", "") or "Color Scheme"
	area_scheme_name = _scheme_area_scheme_name(doc, scheme)
	if area_scheme_name:
		return "Area({}):{}".format(area_scheme_name, scheme_name)
	try:
		cat = doc.GetElement(scheme.CategoryId)
		cat_name = getattr(cat, "Name", "") if cat else ""
	except Exception:
		cat_name = ""
	if not cat_name:
		try:
			cat_name = DB.LabelUtils.GetLabelFor(scheme.CategoryId)
		except Exception:
			cat_name = ""
	if not cat_name:
		try:
			cat_id = scheme.CategoryId
			cat_int = cat_id.IntegerValue if hasattr(cat_id, "IntegerValue") else int(cat_id)
			try:
				import System
				bic = System.Enum.ToObject(DB.BuiltInCategory, cat_int)
				cat_name = DB.LabelUtils.GetLabelFor(bic)
			except Exception:
				cat_name = "Category {}".format(cat_int)
		except Exception:
			cat_name = "Unknown Category"
	return "{}: {}".format(cat_name, scheme_name)


class _SchemeOption(object):
	def __init__(self, name, scheme):
		self.name = name
		self.scheme = scheme


def _copy_scheme_data(source, target, allow_category_mismatch=False):
	try:
		if hasattr(source, "CategoryId") and hasattr(target, "CategoryId"):
			if source.CategoryId != target.CategoryId:
				if not allow_category_mismatch:
					return False, "Source and target schemes use different categories."
	except Exception:
		pass

	for attr in ("Title", "IsByRange", "IsByValue", "IsByPercentage"):
		if hasattr(source, attr) and hasattr(target, attr):
			try:
				setattr(target, attr, getattr(source, attr))
			except Exception:
				pass

	try:
		source_entries = list(source.GetEntries())
	except Exception:
		return False, "Unable to read entries from the source scheme."

	if _call_if_exists(target, "ClearEntries") is None:
		try:
			for entry in list(target.GetEntries()):
				_call_if_exists(target, "RemoveEntry", entry)
		except Exception:
			pass

	try:
		if List is not None:
			entry_list = List[DB.ColorFillSchemeEntry](source_entries)
		else:
			entry_list = list(source_entries)
		if _call_if_exists(target, "SetEntries", entry_list) is not None:
			return True, None
	except Exception:
		pass

	for entry in source_entries:
		try:
			new_entry = entry
			clone = getattr(entry, "Clone", None)
			if callable(clone):
				try:
					new_entry = clone()
				except Exception:
					new_entry = entry
			_call_if_exists(target, "AddEntry", new_entry)
		except Exception:
			pass

	return True, None

doc = revit.doc

schemes = _collect_color_fill_schemes(doc)
if not schemes:
	ui.uiUtils_alert("No Color Fill Schemes found in this model.", title="Copy Color Scheme")
	raise SystemExit

source_names = [_scheme_display_name(doc, s) for s in schemes]
source_indices = ui.uiUtils_select_indices(
	source_names,
	title="Source Color Scheme",
	prompt="Select source scheme:",
	multiselect=False,
	width=720,
	height=520,
)
if not source_indices:
	raise SystemExit
source = schemes[source_indices[0]]

targets = [scheme for scheme in schemes if scheme.Id != source.Id]
if not targets:
	ui.uiUtils_alert("No other Color Fill Scheme found to copy into.", title="Copy Color Scheme")
	raise SystemExit

category_map = {}
for scheme in targets:
	try:
		cat = doc.GetElement(scheme.CategoryId)
		cat_name = getattr(cat, "Name", "") if cat else ""
	except Exception:
		cat_name = ""
	if not cat_name:
		try:
			cat_name = DB.LabelUtils.GetLabelFor(scheme.CategoryId)
		except Exception:
			cat_name = ""
	if not cat_name:
		try:
			cat_id = scheme.CategoryId
			cat_int = cat_id.IntegerValue if hasattr(cat_id, "IntegerValue") else int(cat_id)
			try:
				import System
				bic = System.Enum.ToObject(DB.BuiltInCategory, cat_int)
				cat_name = DB.LabelUtils.GetLabelFor(bic)
			except Exception:
				cat_name = "Category {}".format(cat_int)
		except Exception:
			cat_name = "Unknown Category"
	category_map[cat_name] = scheme.CategoryId

category_names = sorted(category_map.keys(), key=lambda s: (s or "").lower())
selected_indices, overwrite_mode = ui.uiUtils_select_items_with_mode(
	category_names,
	title="Target Category",
	prompt="Select target category:",
	mode_labels=("Overwrite existing", ""),
	default_mode=0,
	prechecked_indices=None,
	width=720,
	height=520,
)
if not selected_indices:
	raise SystemExit

target_category_name = category_names[selected_indices[0]]
target_category_id = category_map.get(target_category_name)
overwrite = overwrite_mode == 1

source_name = getattr(source, "Name", "") or "Color Scheme"
target = None
for scheme in schemes:
	try:
		if scheme.Id == source.Id:
			continue
		if getattr(scheme, "Name", "") != source_name:
			continue
		if scheme.CategoryId != target_category_id:
			continue
		target = scheme
		break
	except Exception:
		continue

if target is None:
	ui.uiUtils_alert(
		"No target Color Scheme named '{}' found in category '{}'.\nCreate a scheme with that name first."
		.format(source_name, target_category_name),
		title="Copy Color Scheme",
	)
	raise SystemExit

if not overwrite:
	ui.uiUtils_alert(
		"Overwrite disabled. Selected target scheme will not be modified.",
		title="Copy Color Scheme",
	)
	raise SystemExit

with revit.Transaction("Copy Color Scheme"):
	try:
		ok, error = _copy_scheme_data(source, target, allow_category_mismatch=True)
		if not ok:
			raise Exception(error or "Copy failed")
	except Exception as ex:
		error = str(ex)
		target = None

if not target:
	ui.uiUtils_alert("Failed to update target Color Scheme.\n\n{}".format(error or "Unknown error"), title="Copy Color Scheme")
	raise SystemExit

try:
	revit.get_selection().set_to([target.Id])
except Exception:
	pass

ui.uiUtils_alert("Updated scheme: {}".format(getattr(target, "Name", "Color Scheme")), title="Copy Color Scheme")
