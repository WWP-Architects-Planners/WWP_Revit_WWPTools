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


def _try_set_element_name(element, new_name):
	# Many Revit elements allow setting .Name directly
	try:
		element.Name = new_name
		return True
	except Exception:
		pass

	# Fallback: try common name parameters
	for bip in (
		getattr(DB.BuiltInParameter, "SYMBOL_NAME_PARAM", None),
		getattr(DB.BuiltInParameter, "ALL_MODEL_TYPE_NAME", None),
		getattr(DB.BuiltInParameter, "DATUM_TEXT", None),
	):
		if bip is None:
			continue
		try:
			p = element.get_Parameter(bip)
			if p and (not p.IsReadOnly):
				p.Set(new_name)
				return True
		except Exception:
			pass

	try:
		p = element.LookupParameter("Name")
		if p and (not p.IsReadOnly):
			p.Set(new_name)
			return True
	except Exception:
		pass

	return False

doc = revit.doc

schemes = _collect_color_fill_schemes(doc)
if not schemes:
	forms.alert("No Color Fill Schemes found in this model.", title="Copy Color Scheme")
	raise SystemExit

source = forms.SelectFromList.show(
	schemes,
	name_attr="Name",
	multiselect=False,
	title="Source Color Scheme",
	button_name="Select Source",
)
if not source:
	raise SystemExit

existing_names = set((getattr(s, "Name", "") or "").lower() for s in schemes)
default_name = "{} (Copy)".format(getattr(source, "Name", "Color Scheme"))
desired_name = forms.ask_for_string(
	default=default_name,
	prompt="New Color Scheme name",
	title="Copy Color Scheme",
)
if desired_name is None:
	raise SystemExit

new_name = _unique_scheme_name(existing_names, desired_name)

new_scheme = None
error = None

with revit.Transaction("Copy Color Scheme"):
	try:
		# Most robust: duplicate the element inside the same document
		new_ids = DB.ElementTransformUtils.CopyElement(doc, source.Id, DB.XYZ(0, 0, 0))
		new_id = list(new_ids)[0]
		new_scheme = doc.GetElement(new_id)
		# Ensure name is set/unique
		if not _try_set_element_name(new_scheme, new_name):
			raise Exception("Created scheme but could not rename it")
	except Exception as ex:
		error = str(ex)

if not new_scheme:
	forms.alert("Failed to create new Color Scheme.\n\n{}".format(error or "Unknown error"), title="Copy Color Scheme")
	raise SystemExit

try:
	# Select the new scheme for convenience
	revit.get_selection().set_to([new_scheme.Id])
except Exception:
	pass

forms.alert("Created new scheme: {}".format(new_name), title="Copy Color Scheme")
