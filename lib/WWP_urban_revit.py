"""Atomic native Revit imports of validated QuickOSM Urban packages."""
import json
import math
import re
import System
from System.Collections.Generic import List
from Autodesk.Revit import DB
from Autodesk.Revit.DB.ExtensibleStorage import Schema,SchemaBuilder,Entity,AccessLevel,DataStorage
from WWP_urban_contract import validate_package,local_to_revit,clean_ring

SCHEMA_GUID=System.Guid('b1401fad-bf08-4e5f-b0bf-e3537d9fd091')

def _schema(create=False):
    schema=Schema.Lookup(SCHEMA_GUID)
    if schema is None and create:
        builder=SchemaBuilder(SCHEMA_GUID)
        builder.SetSchemaName('WWPUrbanContext')
        builder.SetReadAccessLevel(AccessLevel.Public)
        builder.SetWriteAccessLevel(AccessLevel.Public)
        builder.AddSimpleField('Payload',System.String)
        schema=builder.Finish()
    return schema

def _read(element,schema):
    if schema is None:return None
    entity=element.GetEntity(schema)
    if not entity.IsValid():return None
    return json.loads(entity.Get[System.String]('Payload'))

def _write(element,schema,payload):
    entity=Entity(schema)
    entity.Set[System.String]('Payload',json.dumps(payload,ensure_ascii=True,allow_nan=False))
    element.SetEntity(entity)

def find_import(doc,dataset_id):
    """Find the package registry, or None; reject ambiguous duplicated registries."""
    schema=_schema();matches=[]
    for element in DB.FilteredElementCollector(doc).OfClass(DataStorage):
        payload=_read(element,schema)
        if payload and payload.get('kind')=='registry' and payload.get('dataset_id')==dataset_id:
            matches.append((element,payload))
    if len(matches)>1:raise ValueError('Multiple Urban registries have the same package identity')
    return matches[0] if matches else None

def _id(value):
    return DB.ElementId(System.Int64(value))

def _safe_name(value):
    return re.sub(r'[\\:{}\[\]|;<>?`~]',' ',value).strip()[:100]

def _loops(surface,frame,placement,tolerance):
    loops=List[DB.CurveLoop]()
    for ring_index,ring in enumerate(surface['rings']):
        vertices=[]
        for point in clean_ring(ring,tolerance*0.3048)[:-1]:
            value=local_to_revit(point,frame,placement)
            p=DB.XYZ(*value)
            if not vertices or p.DistanceTo(vertices[-1])>tolerance:vertices.append(p)
        while len(vertices)>2 and vertices[0].DistanceTo(vertices[-1])<=tolerance:vertices.pop()
        if len(vertices)<3:raise ValueError('Ring below Revit tolerance in surface '+surface['id'])
        area=sum(vertices[i].X*vertices[(i+1)%len(vertices)].Y-vertices[(i+1)%len(vertices)].X*vertices[i].Y for i in range(len(vertices)))
        if (area<0 and ring_index==0) or (area>0 and ring_index>0):vertices.reverse()
        loop=DB.CurveLoop()
        for i,p in enumerate(vertices):loop.Append(DB.Line.CreateBound(p,vertices[(i+1)%len(vertices)]))
        loops.Add(loop)
    if not DB.BoundaryValidation.IsValidHorizontalBoundary(loops):
        raise ValueError('Revit rejected the boundary of surface '+surface['id']+'; refine the export extent/geometry in QGIS')
    return loops

def _asset(doc,cls,schema,dataset_id,kind,key):
    for element in DB.FilteredElementCollector(doc).OfClass(cls):
        payload=_read(element,schema)
        if payload and payload.get('dataset_id')==dataset_id and payload.get('kind')==kind and payload.get('asset_key')==key:
            return element
    return None

def _element_name(element):
    param=element.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    return param.AsString() if param else element.Name

def _unique_name(doc,cls,name):
    existing=set(_element_name(e) for e in DB.FilteredElementCollector(doc).OfClass(cls))
    candidate=name;index=2
    while candidate in existing:
        candidate=name+' '+str(index);index+=1
    return candidate

def _material(doc,category,solid,schema,dataset_id):
    color=category['rgb'];hex_color=''.join('{:02X}'.format(c) for c in color)
    import hashlib
    suffix=hashlib.sha256((dataset_id+'|'+category['id']).encode('utf-8')).hexdigest()[:8]
    base='WWP Urban - '+_safe_name(category['name'])+' - '+hex_color+' - '+suffix
    material=_asset(doc,DB.Material,schema,dataset_id,'material',category['id'])
    if material is None:
        material=doc.GetElement(DB.Material.Create(doc,_unique_name(doc,DB.Material,base)))
        _write(material,schema,{'dataset_id':dataset_id,'kind':'material','asset_key':category['id']})
    material.Color=DB.Color(*color)
    material.Transparency=0
    material.SurfaceForegroundPatternId=solid.Id
    material.SurfaceForegroundPatternColor=DB.Color(*color)
    return material,base

def _material_type(doc,category,base_type,solid,schema,dataset_id):
    material,base=_material(doc,category,solid,schema,dataset_id)
    name=base+' (10 mm)'
    floor_type=_asset(doc,DB.FloorType,schema,dataset_id,'floor_type',category['id'])
    if floor_type is None:
        floor_type=base_type.Duplicate(_unique_name(doc,DB.FloorType,name))
        _write(floor_type,schema,{'dataset_id':dataset_id,'kind':'floor_type','asset_key':category['id']})
    compound=floor_type.GetCompoundStructure()
    if compound is None:raise ValueError('Selected floor template has no compound structure')
    layers=List[DB.CompoundStructureLayer]()
    layers.Add(DB.CompoundStructureLayer(0.01/0.3048,DB.MaterialFunctionAssignment.Structure,material.Id))
    compound.SetLayers(layers)
    floor_type.SetCompoundStructure(compound)
    return floor_type

def _create_mass(doc,loops,height_m,material,label):
    """Extrude one footprint into a native Mass-category DirectShape."""
    height_ft=float(height_m)/0.3048
    if not 0<height_ft<3300:raise ValueError('Invalid building height for '+label)
    solid=DB.GeometryCreationUtilities.CreateExtrusionGeometry(loops,DB.XYZ.BasisZ,height_ft,DB.SolidOptions(material.Id,DB.ElementId.InvalidElementId))
    shape=DB.DirectShape.CreateElement(doc,DB.ElementId(DB.BuiltInCategory.OST_Mass))
    shape.ApplicationId='WWPTools.UrbanContext'
    shape.ApplicationDataId=label[:200]
    geometry=List[DB.GeometryObject]();geometry.Add(solid)
    shape.SetShape(geometry)
    return shape

def import_package(doc,view,package,placement=None):
    """Create/update only this package's owned floors/text in one undoable transaction."""
    validate_package(package)
    if int(doc.Application.VersionNumber)<2024:raise ValueError('QuickOSM Urban requires Revit 2024 or newer')
    if doc.IsReadOnly or doc.IsFamilyDocument:raise ValueError('Open a writable project document')
    if not isinstance(view,DB.ViewPlan) or view.IsTemplate or view.GenLevel is None:raise ValueError('Choose a non-template plan with an associated level')
    existing=find_import(doc,package['dataset_id'])
    if existing:
        registry,old=existing
        if old['view_id']!=int(view.Id.Value):raise ValueError('Update this package in its original plan view')
        placement=old['placement']
    else:
        registry=None;old={}
        placement=placement or {'crs':package['frame']['crs'],'source_origin':package['frame']['origin'],'anchor_ft':[0,0,view.GenLevel.Elevation],'rotation_rad':0}
    probe=local_to_revit([0,0],package['frame'],placement)
    if any(math.isnan(float(n)) or math.isinf(float(n)) for n in probe):raise ValueError('Invalid placement')
    prepared=[(surface,_loops(surface,package['frame'],placement,doc.Application.ShortCurveTolerance*1.2)) for surface in package['surfaces']]
    label_points=[local_to_revit(label['point'],package['frame'],placement) for label in package.get('labels',[])]
    schema=_schema()
    old_elements=[]
    for element_id in old.get('element_ids',[]):
        element=doc.GetElement(_id(element_id))
        if element is None:continue
        owner=_read(element,schema)
        if not owner or owner.get('dataset_id')!=package['dataset_id'] or owner.get('kind') not in ('floor','label','mass'):
            raise ValueError('An old Urban element no longer has matching ownership; existing import kept')
        old_elements.append(element)
    base_type=next((t for t in DB.FilteredElementCollector(doc).OfClass(DB.FloorType) if not t.IsFoundationSlab and t.GetCompoundStructure() is not None),None)
    base_text=next(iter(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType)),None)
    if base_type is None or (package.get('labels') and base_text is None):raise ValueError('This project needs a floor type and a text-note type')
    solid=next((p for p in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement) if p.GetFillPattern().IsSolidFill),None)
    if solid is None:raise ValueError('No solid fill pattern is available')
    categories={c['id']:c for c in package['categories']}
    txn=DB.Transaction(doc,'Import QuickOSM Urban: '+package['name'])
    txn.Start()
    try:
        schema=_schema(True)
        mass_categories=set(s['category'] for s in package['surfaces'] if 'height_m' in s)
        types={cid:_material_type(doc,c,base_type,solid,schema,package['dataset_id']) for cid,c in categories.items() if cid not in mass_categories}
        materials={cid:_material(doc,categories[cid],solid,schema,package['dataset_id'])[0] for cid in mass_categories}
        text_type=_asset(doc,DB.TextNoteType,schema,package['dataset_id'],'text_type','road_names')
        if text_type is None and base_text is not None:
            text_type=base_text.Duplicate(_unique_name(doc,DB.TextNoteType,'WWP Urban - Road Names 2.5 mm'))
            _write(text_type,schema,{'dataset_id':package['dataset_id'],'kind':'text_type','asset_key':'road_names'})
            text_type.get_Parameter(DB.BuiltInParameter.TEXT_SIZE).Set(2.5/304.8)
        created=[];floors=[];notes=[];masses=[]
        for surface,loops in prepared:
            category=categories[surface['category']]
            if surface['category'] in mass_categories:
                shape=_create_mass(doc,loops,surface['height_m'],materials[surface['category']],package['dataset_id']+'/'+surface['id'])
                comments=shape.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                if comments is not None:comments.Set('WWP Urban | '+package['name']+' | '+category['name']+' | Height (m): {:.2f}'.format(float(surface['height_m'])))
                _write(shape,schema,{'dataset_id':package['dataset_id'],'kind':'mass','source_id':surface['id']})
                created.append(int(shape.Id.Value));masses.append(int(shape.Id.Value))
                continue
            floor=DB.Floor.Create(doc,loops,types[surface['category']].Id,view.GenLevel.Id)
            offset=floor.get_Parameter(DB.BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
            if offset is not None:offset.Set(placement['anchor_ft'][2]-view.GenLevel.Elevation)
            floor.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set('WWP Urban | '+package['name']+' | '+category['name'])
            _write(floor,schema,{'dataset_id':package['dataset_id'],'kind':'floor','source_id':surface['id']})
            overrides=DB.OverrideGraphicSettings()
            overrides.SetSurfaceForegroundPatternId(solid.Id)
            overrides.SetSurfaceForegroundPatternColor(DB.Color(*category['rgb']))
            line=DB.Color(*category.get('outline_rgb',category['rgb']))
            overrides.SetProjectionLineColor(line);overrides.SetCutLineColor(line)
            overrides.SetProjectionLineWeight(1);overrides.SetCutLineWeight(1)
            view.SetElementOverrides(floor.Id,overrides)
            created.append(int(floor.Id.Value));floors.append(int(floor.Id.Value))
        for label,point in zip(package.get('labels',[]),label_points):
            options=DB.TextNoteOptions(text_type.Id)
            options.Rotation=label.get('rotation_rad',0)+placement.get('rotation_rad',0)
            options.HorizontalAlignment=DB.HorizontalTextAlignment.Center
            options.KeepRotatedTextReadable=True
            note=DB.TextNote.Create(doc,view.Id,DB.XYZ(*point),label['text'],options)
            _write(note,schema,{'dataset_id':package['dataset_id'],'kind':'label','source_id':label['id']})
            created.append(int(note.Id.Value));notes.append(int(note.Id.Value))
        old_ids=List[DB.ElementId]()
        for element in old_elements:old_ids.Add(element.Id)
        if old_ids.Count:
            if 'native_dependency_ids' not in old:raise ValueError('This older import lacks dependency tracking; existing context was kept')
            deleted=doc.Delete(old_ids)
            permitted=set(old['native_dependency_ids'])|set(old.get('element_ids',[]))
            unexpected=[int(eid.Value) for eid in deleted if int(eid.Value) not in permitted]
            if unexpected:
                raise ValueError('Updating this context would delete user-added dependent annotations/elements (IDs: '+', '.join(str(i) for i in unexpected[:10])+'). Remove or rehost those references first. Existing context was kept.')
        if registry is None:
            registry=DataStorage.Create(doc)
            registry.Name='WWP Urban '+package['dataset_id']
        doc.Regenerate()
        native_dependencies=set(created)
        for element_id in created:
            for dependent_id in doc.GetElement(_id(element_id)).GetDependentElements(None):native_dependencies.add(int(dependent_id.Value))
        _write(registry,schema,{'kind':'registry','dataset_id':package['dataset_id'],'name':package['name'],'view_id':int(view.Id.Value),'placement':placement,'element_ids':created,'native_dependency_ids':sorted(native_dependencies),'frame':package['frame'],'included_categories':[c['id'] for c in package['categories']]})
        area=sum(doc.GetElement(_id(i)).get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED).AsDouble()*0.3048**2 for i in floors)
        status=txn.Commit()
        if status!=DB.TransactionStatus.Committed:raise ValueError('Revit did not commit the Urban import; existing context was kept')
    except Exception:
        if txn.GetStatus()==DB.TransactionStatus.Started:txn.RollBack()
        raise
    return {'dataset_id':package['dataset_id'],'floors':len(floors),'masses':len(masses),'mass_ids':masses,'floor_ids':floors,'text_notes':len(notes),'text_ids':notes,'floor_types':len(types),'replaced_elements':len(old_elements),'area_m2':area,'registry_id':int(registry.Id.Value),'placement':placement,'crs':package['frame']['crs'],'view_id':int(view.Id.Value)}
