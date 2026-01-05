# Prepare for converters
from pyrevit import revit,DB

# get document
doc = revit.doc


def _conversion_length_unit_type_id():
    return doc.GetUnits().GetFormatOptions(DB.SpecTypeId.Length).GetUnitTypeId()

# Return internal unit type
def conversion_getUnits():
	return _conversion_length_unit_type_id()

# Convert project units to internal
def conversion_prjToInt(length):
	intUnitsId = _conversion_length_unit_type_id()
	return DB.UnitUtils.ConvertToInternalUnits(length, intUnitsId)

# Convert project units from internal
def conversion_intToPrj(length):
	intUnitsId = _conversion_length_unit_type_id()
	return DB.UnitUtils.ConvertFromInternalUnits(length, intUnitsId)
