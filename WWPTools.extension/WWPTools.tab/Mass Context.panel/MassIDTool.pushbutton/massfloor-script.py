#! python3
# pyrevit_engine = CPython312
"""
Copy instance string parameters starting with '*' from Mass to hosted MassFloors
"""
from pyrevit import revit, DB, script
import sys

doc = revit.doc
logger = script.get_logger()
output = script.get_output()

# Provide a `flush()` method for pyRevit's ScriptIO when running under CPython
class _StdWrapper(object):
    def __init__(self, io):
        self._io = io
    def write(self, s):
        return self._io.write(s)
    def flush(self):
        return None
    def __getattr__(self, name):
        return getattr(self._io, name)

sys.stdout = _StdWrapper(sys.stdout)
sys.stderr = _StdWrapper(sys.stderr)


output.print_md("# Mass → MassFloor Parameter Sync")
output.print_md("Copying string instance parameters starting with '*' from Mass to MassFloors...")
output.print_md("---")

# Collect all MassFloor elements
massfloors = list(
    DB.FilteredElementCollector(doc)
    .OfCategory(DB.BuiltInCategory.OST_MassFloor)
    .WhereElementIsNotElementType()
    .ToElements()
)

if not massfloors:
    output.print_md("**No MassFloors found in document.**")
    script.exit()

output.print_md("Found {} MassFloor elements".format(len(massfloors)))

# Start transaction
trans = DB.Transaction(doc, "Copy Mass Parameters to MassFloors")
trans.Start()

try:
    copied_total = 0
    skipped_total = 0
    processed_floors = 0
    no_host_count = 0
    
    # First, let's try to find all the parameters on the first MassFloor to diagnose
    if massfloors:
        output.print_md("### Diagnostics: First MassFloor Parameters")
        first_mf = massfloors[0]
        output.print_md("MassFloor ID: {}".format(first_mf.Id))
        output.print_md("Available BuiltInParameters:")
        for bip in [DB.BuiltInParameter.HOST_ID_PARAM, 
                    DB.BuiltInParameter.SKETCH_PLANE_PARAM,
                    DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM]:
            try:
                p = first_mf.get_Parameter(bip)
                if p and p.HasValue:
                    val = p.AsElementId() if p.StorageType == DB.StorageType.ElementId else p.AsValueString()
                    output.print_md("- {} = {}".format(bip, val))
                else:
                    output.print_md("- {} = (no value)".format(bip))
            except:
                output.print_md("- {} = (not available)".format(bip))
        output.print_md("---")
    
    for mf in massfloors:
        try:
            # Get the host Mass from "Mass Family" parameter
            mass = None
            mass_family_param = mf.LookupParameter("Mass Family")
            
            if mass_family_param and mass_family_param.HasValue:
                if mass_family_param.StorageType == DB.StorageType.ElementId:
                    mass_id = mass_family_param.AsElementId()
                    if mass_id != DB.ElementId.InvalidElementId:
                        mass = doc.GetElement(mass_id)
            
            if not mass:
                no_host_count += 1
                if no_host_count <= 3:  # Only print first 3 to avoid spam
                    output.print_md("MassFloor {} - 'Mass Family' parameter not found or empty".format(mf.Id))
                continue
            
            processed_floors += 1
            output.print_md("### MassFloor {} → Mass {}".format(mf.Id, mass.Id))
            
            # Get all parameters from Mass that start with '*'
            mass_params = {}
            for p in mass.Parameters:
                pname = p.Definition.Name
                if pname.startswith("*"):
                    # Only instance parameters
                    if not p.IsReadOnly:
                        mass_params[pname] = p
            
            if not mass_params:
                output.print_md("_No '*' parameters found on Mass_")
                continue
            
            logger.info("  Found {} '*' parameters on Mass {}".format(len(mass_params), mass.Id))
            
            # Get MassFloor parameters
            floor_params = {}
            for p in mf.Parameters:
                floor_params[p.Definition.Name] = p
            
            copied_this_floor = 0
            skipped_this_floor = 0
            
            for pname, p_mass in mass_params.items():
                # Check if parameter exists on MassFloor
                if pname not in floor_params:
                    logger.debug("  Parameter '{}' not found on MassFloor".format(pname))
                    skipped_this_floor += 1
                    continue
                
                p_floor = floor_params[pname]
                
                # Check if floor parameter is writable
                if p_floor.IsReadOnly:
                    logger.debug("  Parameter '{}' is read-only on MassFloor".format(pname))
                    skipped_this_floor += 1
                    continue
                
                # Check storage type is String
                if p_mass.StorageType != DB.StorageType.String:
                    logger.debug("  Parameter '{}' is not a string (type: {})".format(pname, p_mass.StorageType))
                    skipped_this_floor += 1
                    continue
                
                if p_floor.StorageType != DB.StorageType.String:
                    logger.debug("  Parameter '{}' type mismatch on MassFloor".format(pname))
                    skipped_this_floor += 1
                    continue
                
                # Get value from Mass
                value = p_mass.AsString()
                if not value:
                    logger.debug("  Parameter '{}' is empty on Mass".format(pname))
                    skipped_this_floor += 1
                    continue
                
                # Copy value to MassFloor
                try:
                    p_floor.Set(value)
                    output.print_md("- Copied `{}` = \"{}\"".format(pname, value))
                    copied_this_floor += 1
                    copied_total += 1
                except Exception as set_err:
                    logger.warning("  Failed to set '{}': {}".format(pname, str(set_err)))
                    skipped_this_floor += 1
                    skipped_total += 1
            
            skipped_total += skipped_this_floor
            
            if copied_this_floor == 0:
                output.print_md("_No parameters copied for this floor_")
        
        except Exception as floor_err:
            logger.error("Error processing MassFloor {}: {}".format(mf.Id, str(floor_err)))
            continue
    
    trans.Commit()
    
    # Summary
    output.print_md("---")
    output.print_md("## Summary")
    output.print_md("**Total MassFloors:** {}".format(len(massfloors)))
    output.print_md("**No host found:** {}".format(no_host_count))
    output.print_md("**Processed:** {} MassFloors".format(processed_floors))
    output.print_md("**Copied:** {} parameters".format(copied_total))
    output.print_md("**Skipped:** {} parameters".format(skipped_total))
    output.print_md("---")
    output.print_md("Done.")

except Exception as e:
    trans.RollBack()
    logger.error("Error: {}".format(str(e)))
    output.print_md("**Error:** {}".format(str(e)))
    import traceback
    tb = traceback.format_exc()
    logger.error(tb)
    output.print_md("```text\n{}\n```".format(tb))
