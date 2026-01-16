# -*- coding: utf-8 -*-

from pyrevit import revit, DB, forms

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


def _scheme_display_name(doc, scheme):
	try:
		cat = doc.GetElement(scheme.CategoryId)
		cat_name = getattr(cat, "Name", "") if cat else ""
	except Exception:
		cat_name = ""
	scheme_name = getattr(scheme, "Name", "") or "Color Scheme"
	if cat_name:
		return "{} - {}".format(cat_name, scheme_name)
	return scheme_name


class _SchemeOption(object):
	def __init__(self, name, scheme):
		self.name = name
		self.scheme = scheme


def _copy_scheme_data(source, target):
	try:
		if hasattr(source, "CategoryId") and hasattr(target, "CategoryId"):
			if source.CategoryId != target.CategoryId:
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
	forms.alert("No Color Fill Schemes found in this model.", title="Copy Color Scheme")
	raise SystemExit

source_options = [_SchemeOption(_scheme_display_name(doc, s), s) for s in schemes]
source_option = forms.SelectFromList.show(
	source_options,
	name_attr="name",
	multiselect=False,
	title="Source Color Scheme",
	button_name="Select Source",
)
if not source_option:
	raise SystemExit
source = source_option.scheme

targets = [scheme for scheme in schemes if scheme.Id != source.Id]
if not targets:
	forms.alert("No other Color Fill Scheme found to copy into.", title="Copy Color Scheme")
	raise SystemExit

target_options = [_SchemeOption(_scheme_display_name(doc, s), s) for s in targets]
target_option = forms.SelectFromList.show(
	target_options,
	name_attr="name",
	multiselect=False,
	title="Target Color Scheme",
	button_name="Select Target",
)
if not target_option:
	raise SystemExit
target = target_option.scheme

with revit.Transaction("Copy Color Scheme"):
	try:
		ok, error = _copy_scheme_data(source, target)
		if not ok:
			raise Exception(error or "Copy failed")
	except Exception as ex:
		error = str(ex)
		target = None

if not target:
	forms.alert("Failed to update target Color Scheme.\n\n{}".format(error or "Unknown error"), title="Copy Color Scheme")
	raise SystemExit

try:
	revit.get_selection().set_to([target.Id])
except Exception:
	pass

forms.alert("Updated scheme: {}".format(getattr(target, "Name", "Color Scheme")), title="Copy Color Scheme")
