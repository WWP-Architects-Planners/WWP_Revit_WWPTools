import sys
import os
from collections import namedtuple
import clr

clr.AddReference("RevitNodes")
import Revit

#from Autodesk.Revit.DB.Architecture import Area

import rpw
from rpw import revit, doc, uidoc, DB, UI, db, ui



# Validate + Filter Selection
selection = rpw.ui.Selection()
selected_areas = [e for e in selection.elements]

#if not selected_areas:
#    ui.forms.Alert('MakeViews', 'You need to select at lest one Area.')

# Get View Types and Prompt User
plan_types = db.Collector(of_class='ViewFamilyType', is_type=True).wrapped_elements

# Filter all view types that are FloorPlan or CeilingPlan
plan_types_options = {t.name: t for t in plan_types
                      if t.view_family.name in ('FloorPlan', 'CeilingPlan')}

plan_type = ui.forms.SelectFromList('MakeViews', plan_types_options,
                                    description='Select View Type')
view_type_id = plan_type.Id


@rpw.db.Transaction.ensure('Create View')
def create_plan(new_view, view_type_id, cropbox_visible=False, remove_underlay=True):
    """Create a Drafting View"""

    view_type_id
    name = new_view.name
    bbox = new_view.bbox
    level_id = new_view.level_id
    viewplan = DB.ViewPlan.CreatePlan(doc, view_type_id, level_id)
    viewplan.CropBoxActive = True
    viewplan.CropBoxVisible = cropbox_visible
    if remove_underlay and revit.version.year == '2015':
        underlay_param = viewplan.get_Parameter(DB.BuiltInParameter.VIEW_UNDERLAY_ID)
        underlay_param.Set(DB.ElementId.InvalidElementId)
    viewplan.CropBox = bbox

    counter = 1
    while True:
        # Auto Increment area Number
        try:
            viewplan.Name = name
        except Exception:
            try:
                viewplan.Name = '{} - Copy {}'.format(name, counter)
            except Exception as errmsg:
                counter += 1
                if counter > 100:
                    raise Exception('Exceeded Maximum Loop')
            else:
                break
        else:
            break
    return viewplan

NewView = namedtuple('NewView', ['name', 'bbox', 'level_id'])
new_views = []

for area in selected_areas:
        area = db.Element(area)
        area_level_id = area.Level.Id
        area_name = area.parameters['Name'].value
        area_number = area.parameters['Number'].value

        new_area_name = '{} {}'.format(area_name, area_number)
        area_bbox = area.get_BoundingBox(doc.ActiveView)
        view_name = '{} - {}'.format(area.Level.Name, new_area_name)
        new_view = NewView(name=view_name, bbox=area_bbox, level_id=area_level_id)
        new_views.append(new_view)

for new_view in new_views:
    view = create_plan(new_view= new_view, view_type_id=view_type_id)