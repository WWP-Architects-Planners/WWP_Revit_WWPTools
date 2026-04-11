"""Purge unused DirectShape type definitions left by Data Exchange or other connectors."""

from System.Collections.Generic import List
from pyrevit import revit, DB
from WWP_uiUtils import uiUtils_alert, uiUtils_confirm, uiUtils_select_indices

TITLE = "Purge Data Exchange Geometry"

doc = revit.doc


# Collect type IDs that are actively referenced by placed DirectShape instances
used_type_ids = set()
for shape in DB.FilteredElementCollector(doc).OfClass(DB.DirectShape):
    type_id = shape.GetTypeId()
    if type_id and type_id != DB.ElementId.InvalidElementId:
        used_type_ids.add(type_id.Value)

# All DirectShape type definitions in the document
all_types = list(DB.FilteredElementCollector(doc).OfClass(DB.DirectShapeType))

# Keep only the ones with no placed instances (i.e. what Revit's Purge Unused would catch)
unused_types = [t for t in all_types if t.Id.Value not in used_type_ids]

if not unused_types:
    uiUtils_alert("No unused DirectShape types found in this document.", TITLE)
else:
    def _element_name(t):
        # Revit 2024+ may not expose .Name directly on DirectShapeType
        try:
            n = t.Name
            if n:
                return n
        except Exception:
            pass
        for bip in (DB.BuiltInParameter.ALL_MODEL_TYPE_NAME, DB.BuiltInParameter.SYMBOL_NAME_PARAM):
            try:
                param = t.get_Parameter(bip)
                if param:
                    val = param.AsString()
                    if val:
                        return val
            except Exception:
                pass
        return "<unnamed>"

    def _element_category(t):
        try:
            return t.Category.Name
        except Exception:
            pass
        try:
            return str(t.Category.Id)
        except Exception:
            return "Unknown"

    def _type_display(t):
        return "{} | {}".format(_element_category(t), _element_name(t))

    display = [_type_display(t) for t in unused_types]

    indices = uiUtils_select_indices(
        display,
        title=TITLE,
        prompt="Select unused DirectShape types to purge ({} found):".format(len(unused_types))
    )

    selected = [unused_types[i] for i in (indices or [])]

    if selected and uiUtils_confirm(
        "Purge {} unused DirectShape type(s)?\n\nThis cannot be undone.".format(len(selected)),
        TITLE
    ):
        removed = []
        failed = []

        for t in selected:
            with revit.Transaction("Purge Data Exchange Geometry"):
                try:
                    doc.Delete(t.Id)
                    removed.append(t)
                except Exception as e:
                    failed.append((_type_display(t), str(e)))

        if failed:
            msg = "Purged {}. Failed {}:\n\n".format(len(removed), len(failed))
            msg += "\n".join("- {} | {}".format(n, err) for n, err in failed[:20])
            uiUtils_alert(msg, TITLE)
        else:
            uiUtils_alert("Purged {} DirectShape type(s).".format(len(removed)), TITLE)
