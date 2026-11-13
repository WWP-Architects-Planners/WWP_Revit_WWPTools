"""Import QuickOSM Urban roads and road names. Shift-click installs the QGIS exporter."""
import io
import math
import os
import shutil
import System
from pyrevit import revit,DB,forms
from Autodesk.Revit.Exceptions import OperationCanceledException
import WWP_uiUtils as ui
from WWP_settings import ProjectToolSettings
from WWP_urban_contract import load_package,filter_categories
from WWP_urban_revit import find_import,import_package

TITLE='QuickOSM Urban Context'
BUNDLE=os.path.dirname(__file__)

def _initial_directory(settings):
    try:return os.path.dirname(getattr(settings,'last_package','') or '')
    except (TypeError,ValueError):return ''

def _remember_package(settings,path):
    try:
        settings.last_package=path
        settings.save()
        return None
    except Exception as exc:
        return str(exc)

def _width_description(package):
    if package.get('width_mode')=='qgis_style':
        return 'QGIS road-body widths frozen at 1:{:g} (visual context)'.format(float(package.get('width_reference_scale',package.get('qgis_scale',2500))))
    return 'Physical / estimated widths (OSM width, lanes, then road class)'

class RoadGroupOption(forms.TemplateListItem):
    @property
    def name(self):
        color=''.join('{:02X}'.format(c) for c in self.item['rgb'])
        return '{}  |  {} {}  |  #{}'.format(self.item['name'],self.item['surface_count'],'mass buildings' if self.item.get('is_mass') else 'floor surfaces',color)

def _select_road_groups(package,existing):
    counts={};mass_ids=set(s['category'] for s in package['surfaces'] if 'height_m' in s)
    for surface in package['surfaces']:counts[surface['category']]=counts.get(surface['category'],0)+1
    defaults=set(existing[1].get('included_categories',counts)) if existing else set(counts)
    if not defaults.intersection(counts):defaults=set(counts)
    options=[]
    for category in sorted(package['categories'],key=lambda c:c['name']):
        if category['id'] not in counts:continue
        item=dict(category);item['surface_count']=counts[category['id']];item['is_mass']=category['id'] in mass_ids
        options.append(RoadGroupOption(item,checked=category['id'] in defaults))
    selected=forms.SelectFromList.show(options,title='Choose road groups to import - uncheck groups to omit',multiselect=True,button_name='Use checked road groups',width=900,height=620)
    if not selected:return None
    return filter_categories(package,[c['id'] for c in selected])

def install_exporter(folder):
    """Install the Processing algorithm and its private shared package validator."""
    if not os.path.isabs(folder):raise ValueError('Choose an absolute QGIS scripts folder')
    library=os.path.join(folder,'wwp_urban')
    if not os.path.isdir(library):os.makedirs(library)
    import WWP_urban_contract
    shutil.copy2(WWP_urban_contract.__file__.replace('.pyc','.py'),os.path.join(library,'WWP_urban_contract.py'))
    destination=os.path.join(folder,'ExportQuickOSMUrban.py')
    shutil.copy2(os.path.join(BUNDLE,'qgis','ExportQuickOSMUrban.py'),destination)
    return destination

def main():
    if globals().get('__shiftclick__',False):
        default=os.path.join(os.getenv('APPDATA'),'QGIS','QGIS3','profiles','default','processing','scripts')
        indices=ui.uiUtils_select_indices(['Install to default QGIS profile','Choose a different Processing scripts folder'],title=TITLE,prompt='Install the QGIS exporter',multiselect=False,width=640,height=340)
        if not indices:return
        folder=default if indices[0]==0 else ui.uiUtils_select_folder_dialog(title='QGIS Processing scripts folder',initial_directory=os.path.dirname(default))
        if not folder:return
        destination=install_exporter(folder)
        ui.uiUtils_alert('Installed:\n'+destination+'\n\nIn QGIS, refresh the Processing Scripts provider or restart QGIS. Search for "Export QuickOSM Urban for Revit".',TITLE)
        return
    doc=revit.doc
    settings=ProjectToolSettings('QuickOSMUrban',doc)
    path=forms.pick_file(file_ext='json',init_dir=_initial_directory(settings),title='Choose an exported Urban context package')
    if not path:return
    package=load_package(path)
    existing=find_import(doc,package['dataset_id'])
    package=_select_road_groups(package,existing)
    if package is None:return
    placement=None
    if existing:
        view=doc.GetElement(DB.ElementId(System.Int64(existing[1]['view_id'])))
        if view is None:raise ValueError('The original import view was deleted. Restore it or export to a new file for a new import.')
    else:
        view=doc.ActiveView
        if not isinstance(view,DB.ViewPlan) or view.IsTemplate:
            views=[v for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewPlan) if not v.IsTemplate and v.GenLevel is not None]
            views.sort(key=lambda v:v.Name)
            indices=ui.uiUtils_select_indices([v.Name for v in views],title=TITLE,prompt='Choose a plan for the roads and text',multiselect=False)
            if not indices:return
            view=views[indices[0]]
        indices=ui.uiUtils_select_indices(['Place export origin at Revit internal origin','Pick export origin in the plan'],title=TITLE,prompt='Place the package origin (shared coordinates are not changed)',multiselect=False,width=740,height=350)
        if not indices:return
        anchor=[0,0,view.GenLevel.Elevation]
        if indices[0]==1:
            if doc.ActiveView.Id!=view.Id:revit.uidoc.ActiveView=view
            try:
                picked=revit.uidoc.Selection.PickPoint('Pick the Revit location of the QGIS export origin')
            except OperationCanceledException:return
            anchor=[picked.X,picked.Y,view.GenLevel.Elevation]
        rotation=forms.ask_for_string(default='0',prompt='Rotation counterclockwise from Project North, in degrees. Zero keeps the exported grid north up.',title=TITLE)
        if rotation is None:return
        angle=float(rotation)
        if math.isnan(angle) or math.isinf(angle):raise ValueError('Enter a finite rotation angle')
        placement={'crs':package['frame']['crs'],'source_origin':package['frame']['origin'],'anchor_ft':anchor,'rotation_rad':math.radians(angle)}
    frame=package['frame']
    report='{}\n\nExport CRS: {}\nSource CRS: {}\nUnits: metres\nExport origin: {}, {}\nTarget plan: {}\n\n{} surfaces, {} categories, {} text notes.\nWidth mode: {}\n\n{}\n\nRoads are flat floors; buildings are Mass-category extrusions using each footprint height. Shared coordinates are not changed.'.format(package['name'],frame['crs'],frame.get('source_crs','recorded by exporter'),frame['origin'][0],frame['origin'][1],view.Name,len(package['surfaces']),len(package['categories']),len(package.get('labels',[])),_width_description(package),'Updates only this package\'s previously imported elements.' if existing else 'Creates a new context import.')
    if not ui.uiUtils_show_text_report(TITLE,report,ok_text='Update Context' if existing else 'Import Context',cancel_text='Cancel',width=760,height=650):return
    result=import_package(doc,view,package,placement)
    settings_error=_remember_package(settings,path)
    message='Imported {} floors across {} types, {} Mass-category buildings and {} text notes into {}.\n\nReplaced {} elements from this package.\nExport CRS: {}\n\nExport to the same JSON file and import again to update it.'.format(result['floors'],result['floor_types'],result['masses'],result['text_notes'],view.Name,result['replaced_elements'],result['crs'])
    if settings_error:message+='\n\nContext was created, but the last package path could not be saved: '+settings_error
    ui.uiUtils_alert(message,TITLE)

if __name__=='__main__':
    try:main()
    except Exception as exc:
        ui.uiUtils_alert('Urban context was not imported.\n\n'+str(exc),TITLE)
