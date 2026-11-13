"""Export QuickOSM Urban roads for WWPTools. QGIS 3.34 or newer."""
import hashlib
import json
import math
import os
import re
import sys
from datetime import datetime
from qgis.core import (QgsCoordinateReferenceSystem,QgsCoordinateTransform,QgsDistanceArea,QgsFeatureRequest,QgsGeometry,QgsMapToPixel,QgsPointXY,QgsProcessing,QgsProcessingAlgorithm,QgsProcessingException,QgsProcessingFeedback,QgsProcessingParameterCrs,QgsProcessingParameterEnum,QgsProcessingParameterExtent,QgsProcessingParameterField,QgsProcessingParameterFileDestination,QgsProcessingParameterNumber,QgsProcessingParameterString,QgsProcessingParameterVectorLayer,QgsProject,QgsRectangle,QgsRenderContext,QgsRuleBasedRenderer,QgsSimpleLineSymbolLayer,QgsUnitTypes,QgsWkbTypes)

_folder=os.path.dirname(os.path.abspath(__file__))
_search=[os.path.join(_folder,'wwp_urban')]
_parent=_folder
while True:
    _search.append(os.path.join(_parent,'lib'))
    _next=os.path.dirname(_parent)
    if _next==_parent:break
    _parent=_next
for _candidate in _search:
    if os.path.isfile(os.path.join(_candidate,'WWP_urban_contract.py')):
        sys.path.insert(0,_candidate)
        break
from WWP_urban_contract import FORMAT,VERSION,write_package

DEFAULT_WIDTHS={'motorway':14,'motorway_link':6,'trunk':14,'trunk_link':6,'primary':12,'primary_link':6,'secondary':10,'secondary_link':6,'tertiary':8,'tertiary_link':6,'residential':7.2,'service':5,'living_street':5,'unclassified':6.5,'pedestrian':4,'cycleway':3,'path':2,'footway':2,'footpath':2,'steps':2,'track':4}

def _string(value):
    return '' if value is None or str(value)=='NULL' else str(value)

def _tags(feature):
    tags={f.name():_string(feature[f.name()]) for f in feature.fields()}
    for key,value in re.findall(r'"([^"\\]+)"=>"((?:\\.|[^"\\])*)"',tags.get('other_tags','')):
        if not tags.get(key):
            try:tags[key]=json.loads('"'+value+'"')
            except ValueError:tags[key]=value
    return tags

def road_width(tags):
    for field in ('width','est_width'):
        match=re.fullmatch(r'\s*([0-9]+(?:\.[0-9]+)?)\s*(m|metres|meters|ft|feet|\x27)?\s*',tags.get(field,''),re.I)
        if match:
            width=float(match.group(1))*(0.3048 if (match.group(2) or '').lower() in ('ft','feet',"'") else 1)
            if 0.5<=width<=60:return width,'OSM '+field
    highway=tags.get('highway','')
    lanes=tags.get('lanes','')
    if lanes.isdigit() and 1<=int(lanes)<=12 and highway not in ('service','cycleway','path','footpath','footway','pedestrian','steps','track'):
        return int(lanes)*3.2,'estimated from lanes at 3.2 m/lane'
    return DEFAULT_WIDTHS.get(highway,7),'estimated by highway class'

BUILDING_ID='buildings'
BUILDING_CATEGORY={'id':BUILDING_ID,'name':'Buildings','rgb':[200,200,200],'outline_rgb':[120,120,120],'priority':-1}

def _metres(text):
    match=re.fullmatch(r'\s*([0-9]+(?:\.[0-9]+)?)\s*(m|metres|meters|ft|feet|\x27)?\s*',text or '',re.I)
    if not match:return None
    value=float(match.group(1))*(0.3048 if (match.group(2) or '').lower() in ('ft','feet',"'") else 1)
    return value if 0.5<=value<=1000 else None

def building_height(tags,field='',key=''):
    """Height in metres: chosen column, then OSM height, then building:levels x 3 m, else a stable 6-8 m."""
    for name in ([field] if field else [])+['height']:
        value=_metres(tags.get(name,''))
        if value is not None:return value,'column '+name if name==field else 'OSM height'
    levels=tags.get('building:levels','')
    if re.fullmatch(r'\s*[0-9]+(?:\.[0-9]+)?\s*',levels) and 0<float(levels)<=300:return float(levels)*3.0,'estimated from building:levels at 3 m/level'
    fraction=int(hashlib.sha256(str(key).encode('utf-8')).hexdigest()[:8],16)/float(0xffffffff)
    return round(6.0+2.0*fraction,1),'no height data: stable pseudo-random 6-8 m'

def _check_cancel(feedback):
    if feedback.isCanceled():raise QgsProcessingException('Export cancelled; existing package kept')

def _crs_identity(crs):
    authority=crs.authid()
    if authority and not authority.upper().startswith(('USER:','INTERNAL:')):return authority
    return 'WKT:'+hashlib.sha256(crs.toWkt().encode('utf-8')).hexdigest()

def _parts(geometry,kind):
    if geometry.isEmpty():return []
    if geometry.isMultipart() or QgsWkbTypes.flatType(geometry.wkbType())==QgsWkbTypes.GeometryCollection:
        result=[]
        for part in geometry.asGeometryCollection():result.extend(_parts(part,kind))
        return result
    return [geometry] if geometry.type()==kind else []

def _symbol_layers(rule):
    symbol=rule.symbol()
    layers=[symbol.symbolLayer(i) for i in range(symbol.symbolLayerCount()) if symbol.symbolLayer(i).enabled()]
    if not layers:raise QgsProcessingException('The road rule has no enabled symbol layers: '+rule.label())
    solid=[layer for layer in layers if layer.properties().get('line_style','solid')=='solid']
    if not solid:solid=layers
    return layers,solid

def _symbol_category(rule):
    layers,solid=_symbol_layers(rule)
    body=solid[-1].color();outline=solid[0].color()
    name=rule.label() or rule.filterExpression() or 'Road'
    key=hashlib.sha256((name+'|'+rule.filterExpression()+'|'+body.name()).encode('utf-8')).hexdigest()[:16]
    return {'id':key,'name':name,'rgb':[body.red(),body.green(),body.blue()],'outline_rgb':[outline.red(),outline.green(),outline.blue()],'priority':max(layer.renderingPass() for layer in layers),'qgis_filter':rule.filterExpression()}

def _style_width(rule,render):
    layers,solid=_symbol_layers(rule)
    body=solid[-1]
    if not isinstance(body,QgsSimpleLineSymbolLayer) or body.dataDefinedProperties().hasActiveProperties():
        raise QgsProcessingException('Match QGIS style requires a fixed SimpleLine road body; use physical widths or simplify the style for '+rule.label())
    width=render.convertToMapUnits(body.width(),body.widthUnit(),body.widthMapUnitScale())
    if not math.isfinite(width) or not 0.01<=width<=500:
        raise QgsProcessingException('Styled road width is outside 0.01 to 500 metres. Choose a suitable reference scale for '+rule.label())
    return width,'QGIS style road body at 1:%g (visual width, not surveyed)'%render.rendererScale()

def export_layer(layer,extent,extent_crs,scale,output_path,feedback=None,dataset_name='Urban context',target_crs=None,width_mode='physical',buildings=None,height_field=''):
    """Export one road layer (plus an optional whole building polygon layer) atomically; coordinates use projected metres."""
    feedback=feedback or QgsProcessingFeedback()
    _check_cancel(feedback)
    if width_mode not in ('physical','qgis_style'):raise QgsProcessingException('Choose a supported road width mode')
    if not math.isfinite(float(scale)) or float(scale)<1:raise QgsProcessingException('Choose a finite map scale denominator of at least 1')
    if layer is None or not layer.isValid() or layer.geometryType()!=QgsWkbTypes.LineGeometry:
        raise QgsProcessingException('Choose the QuickOSM Urban road line layer')
    if not layer.crs().isValid():raise QgsProcessingException('The source road layer has no CRS. Identify its real source CRS in QGIS before exporting; do not guess from the canvas CRS.')
    if layer.fields().indexOf('highway')<0:
        raise QgsProcessingException('The selected layer has no highway field; choose the Urban road layer')
    if not isinstance(layer.renderer(),QgsRuleBasedRenderer):
        raise QgsProcessingException('Use the rule-based QuickOSM Urban road style')
    context=QgsProject.instance().transformContext()
    full_layer=extent is None
    if full_layer:
        extent=layer.extent();extent_crs=layer.crs()
    extent_crs=extent_crs or layer.crs()
    if extent.isEmpty():raise QgsProcessingException('Choose a nonempty export extent')
    wgs=QgsCoordinateReferenceSystem('EPSG:4326')
    center=QgsCoordinateTransform(extent_crs,wgs,context).transform(extent.center())
    if not -80<=center.y()<=84:raise QgsProcessingException('This exporter requires a location within UTM coverage (-80 to 84 degrees latitude)')
    zone=min(60,max(1,int((center.x()+180)/6)+1))
    projected=QgsCoordinateReferenceSystem('EPSG:'+str((32600 if center.y()>=0 else 32700)+zone))
    if target_crs is not None:projected=QgsCoordinateReferenceSystem(target_crs)
    if not projected.isValid() or projected.isGeographic() or projected.mapUnits()!=QgsUnitTypes.DistanceMeters or projected.authid()=='EPSG:3857':
        raise QgsProcessingException('Choose a local projected CRS measured in metres, such as the project UTM/MTM CRS. Geographic degrees, Web Mercator and feet are not supported for physical road buffers.')
    crs_identity=_crs_identity(projected)
    if os.path.isfile(output_path):
        from WWP_urban_contract import load_package
        previous=load_package(output_path)
        restored=QgsCoordinateReferenceSystem()
        if previous['frame'].get('crs_wkt'):restored.createFromWkt(previous['frame']['crs_wkt'])
        else:restored=QgsCoordinateReferenceSystem(previous['frame']['crs'])
        if target_crs is not None and projected!=restored:
            raise QgsProcessingException('Export CRS differs from the existing package. Keep its CRS or export to a new file for a separate import.')
        projected=restored;crs_identity=previous['frame']['crs']
        if not projected.isValid() or projected.isGeographic():raise QgsProcessingException('Existing package CRS is not usable')
    projection=QgsCoordinateTransform(layer.crs(),projected,context)
    extent_projection=QgsCoordinateTransform(extent_crs,projected,context)
    boundary=QgsGeometry.fromRect(extent)
    boundary.transform(extent_projection)
    margin=250 if width_mode=='qgis_style' else 30
    if full_layer:boundary=boundary.buffer(margin,1)
    origin=boundary.boundingBox().center();ox,oy=origin.x(),origin.y()
    search_metric=QgsRectangle(boundary.boundingBox())
    search_metric.grow(margin)
    search_rect=QgsCoordinateTransform(projected,layer.crs(),context).transformBoundingBox(search_metric)
    render=QgsRenderContext();render.setRendererScale(float(scale))
    render.setScaleFactor(96/25.4)
    render.setMapToPixel(QgsMapToPixel(float(scale)*0.0254/96))
    distances=QgsDistanceArea();distances.setSourceCrs(projected,context)
    render.setDistanceArea(distances)
    renderer=layer.renderer().clone();renderer.startRender(render,layer.fields())
    categories={};groups={};sources=[];named={};name_categories={};count=0
    try:
        for feature in layer.getFeatures(QgsFeatureRequest().setFilterRect(search_rect)):
            _check_cancel(feedback);count+=1
            if count>20000:raise QgsProcessingException('Too many source features; reduce the export extent')
            tags=_tags(feature)
            if not tags.get('highway'):continue
            matches=[r for r in renderer.rootRule().rulesForFeature(feature,render) if r.symbol() is not None]
            if not matches:continue
            rule=max(matches,key=lambda r:max(r.symbol().symbolLayer(i).renderingPass() for i in range(r.symbol().symbolLayerCount())))
            category=_symbol_category(rule);cid=category['id'];categories[cid]=category
            line=QgsGeometry(feature.geometry());line.transform(projection)
            if line.isEmpty():continue
            width,width_source=_style_width(rule,render) if width_mode=='qgis_style' else road_width(tags)
            polygon=line.buffer(width/2,6)
            if not polygon.isGeosValid():polygon=polygon.makeValid()
            polygon=polygon.intersection(boundary)
            if polygon.isEmpty():continue
            if not polygon.isGeosValid():polygon=polygon.makeValid()
            if polygon.isEmpty() or not polygon.isGeosValid():raise QgsProcessingException('Cannot create valid road polygon for feature '+str(feature.id()))
            polygon.translate(-ox,-oy)
            groups.setdefault(cid,[]).append(polygon)
            source_id=tags.get('osm_id') or tags.get('full_id') or str(feature.id())
            sources.append({'id':source_id,'category':cid,'highway':tags['highway'],'name':tags.get('name',''),'width_m':width,'width_source':width_source,'bridge':tags.get('bridge',''),'tunnel':tags.get('tunnel',''),'layer':tags.get('layer','')})
            name=tags.get('name','').strip()
            if name:
                clipped=line.intersection(boundary)
                clipped.translate(-ox,-oy)
                named.setdefault(name,[]).extend(_parts(clipped,QgsWkbTypes.LineGeometry))
                name_categories.setdefault(name,set()).add(cid)
    finally:
        renderer.stopRender(render)
    if not groups:raise QgsProcessingException('No visible Urban roads in this extent at the selected map scale; existing package kept')
    surfaces=[];used=QgsGeometry()
    for category in sorted(categories.values(),key=lambda c:(-c['priority'],c['id'])):
        _check_cancel(feedback)
        cid=category['id']
        if cid not in groups:continue
        merged=QgsGeometry.unaryUnion(groups[cid])
        if merged.isEmpty():raise QgsProcessingException('Road dissolve failed for '+category['name'])
        visible=merged if used.isEmpty() else merged.difference(used)
        used=merged if used.isEmpty() else QgsGeometry.unaryUnion([used,merged])
        if not visible.isGeosValid():visible=visible.makeValid()
        for part in _parts(visible,QgsWkbTypes.PolygonGeometry):
            if part.area()<0.02:continue
            rings=[[[p.x(),p.y()] for p in ring] for ring in part.asPolygon()]
            surfaces.append({'id':cid+'/'+str(len(surfaces)),'category':cid,'rings':rings})
    building_surfaces=[];building_sources=[]
    if buildings is not None:
        if not buildings.isValid() or buildings.geometryType()!=QgsWkbTypes.PolygonGeometry:
            raise QgsProcessingException('Choose the QuickOSM building polygon layer')
        if not buildings.crs().isValid():
            raise QgsProcessingException('The building layer has no CRS. Identify its real source CRS in QGIS before exporting.')
        building_projection=QgsCoordinateTransform(buildings.crs(),projected,context)
        bcount=0
        for feature in buildings.getFeatures():
            _check_cancel(feedback);bcount+=1
            if bcount>20000:raise QgsProcessingException('Too many building features (limit 20000); split or filter the building layer')
            shape=QgsGeometry(feature.geometry())
            if shape.isEmpty():continue
            shape.transform(building_projection)
            if not shape.isGeosValid():shape=shape.makeValid()
            shape.translate(-ox,-oy)
            tags=_tags(feature)
            source_id=tags.get('osm_id') or tags.get('osm_way_id') or tags.get('full_id') or str(feature.id())
            height,height_source=building_height(tags,height_field,source_id)
            added=False
            for part in _parts(shape,QgsWkbTypes.PolygonGeometry):
                if part.area()<0.5:continue
                rings=[[[p.x(),p.y()] for p in ring] for ring in part.asPolygon()]
                building_surfaces.append({'id':BUILDING_ID+'/'+str(len(building_surfaces)),'category':BUILDING_ID,'rings':rings,'height_m':height,'source':source_id});added=True
            if added:building_sources.append({'id':source_id,'category':BUILDING_ID,'building':tags.get('building',''),'name':tags.get('name',''),'height_m':height,'height_source':height_source})
        if not building_surfaces:raise QgsProcessingException('The building layer has no usable polygons; existing package kept')
        surfaces.extend(building_surfaces)
    labels=[];occupied=[]
    for name,lines in sorted(named.items(),key=lambda item:-sum(g.length() for g in item[1])):
        merged=QgsGeometry.unaryUnion(lines).mergeLines()
        parts=_parts(merged,QgsWkbTypes.LineGeometry)
        if not parts:continue
        line=max(parts,key=lambda g:g.length());length=line.length()
        if length<0.01:continue
        distance=length/2;p=line.interpolate(distance).asPoint()
        for fraction in (0.5,0.25,0.75,0.1,0.9):
            candidate=line.interpolate(length*fraction).asPoint()
            if not any(math.hypot(candidate.x()-x,candidate.y()-y)<25 for x,y in occupied):
                distance=length*fraction;p=candidate;break
        a=line.interpolate(max(0,distance-5)).asPoint();b=line.interpolate(min(length,distance+5)).asPoint()
        angle=math.atan2(b.y()-a.y(),b.x()-a.x())
        if angle>math.pi/2:angle-=math.pi
        if angle<-math.pi/2:angle+=math.pi
        labels.append({'id':hashlib.sha256(name.encode('utf-8')).hexdigest()[:20],'text':name,'point':[p.x(),p.y()],'rotation_rad':angle,'categories':sorted(name_categories[name])})
        occupied.append((p.x(),p.y()))
    data={'format':FORMAT,'version':VERSION,'name':dataset_name.strip() or layer.name(),'frame':{'crs':crs_identity,'crs_description':projected.description(),'crs_wkt':projected.toWkt(),'source_crs':_crs_identity(layer.crs()),'source_crs_wkt':layer.crs().toWkt(),'units':'metres','origin':[ox,oy],'north_reference':'projected grid north'},'categories':[categories[cid] for cid in sorted(groups)]+([BUILDING_CATEGORY] if building_surfaces else []),'surfaces':surfaces,'labels':labels,'sources':sources+building_sources,'qgis_scale':float(scale),'exported_utc':datetime.utcnow().isoformat()+'Z','source_layer':layer.name(),'label_placement':'Straight labels derived from road names/geometry; curved map label placement is not reproduced','notes':'Flat context surfaces. Bridge/tunnel/layer tags are retained in sources; elevations and decorative line overlays are not modeled. Data copyright OpenStreetMap contributors, ODbL: https://www.openstreetmap.org/copyright'}
    data['export_scope']='full_layer' if full_layer else 'specified_extent'
    data['width_mode']=width_mode
    data['width_reference_scale']=float(scale) if width_mode=='qgis_style' else None
    data['source_feature_count']=count
    data['source_extent']=[layer.extent().xMinimum(),layer.extent().yMinimum(),layer.extent().xMaximum(),layer.extent().yMaximum()]
    data['export_bounds_local']=list(used.boundingBox().toRectF().getCoords())
    _check_cancel(feedback)
    result=write_package(output_path,data)
    feedback.pushInfo('Exported %d surfaces, %d types and %d road names.'%(len(surfaces),len(data['categories']),len(labels)))
    return result

class ExportQuickOSMUrban(QgsProcessingAlgorithm):
    def name(self):return 'exportquickosmurban'
    def displayName(self):return 'Export QuickOSM Urban for Revit'
    def group(self):return 'WWPTools'
    def groupId(self):return 'wwptools'
    def createInstance(self):return ExportQuickOSMUrban()
    def flags(self):return super().flags()|QgsProcessingAlgorithm.FlagNoThreading
    def shortHelpString(self):return 'Exports the FULL styled Urban road layer by default, independently of screen zoom. Choose Specified extent only when you want a crop. Match QGIS style converts the colored road body width into metres at the chosen reference scale; 1 mm at 1:2500 becomes 2.5 m. Physical/estimated mode uses OSM widths, then lane/class estimates. Colors and straight road-name labels are retained; building polygons are exported whole with height_m when a building layer is chosen. Export to the same JSON file to update the same Revit import.'
    def initAlgorithm(self,config=None):
        default_layer=None;default_extent=None;default_scale=2500;default_crs=None
        try:
            from qgis.utils import iface
            if iface:
                candidates=[l for l in QgsProject.instance().mapLayers().values() if l.name().lower()=='road']
                if candidates:default_layer=candidates[0].id()
                canvas=iface.mapCanvas();rect=canvas.extent();crs=canvas.mapSettings().destinationCrs().authid()
                default_extent='%s,%s,%s,%s [%s]'%(rect.xMinimum(),rect.xMaximum(),rect.yMinimum(),rect.yMaximum(),crs)
                default_scale=canvas.scale()
                project_crs=canvas.mapSettings().destinationCrs()
                if project_crs.isValid() and not project_crs.isGeographic() and project_crs.mapUnits()==QgsUnitTypes.DistanceMeters and project_crs.authid()!='EPSG:3857':
                    default_crs=project_crs.authid()
                else:
                    center=QgsCoordinateTransform(project_crs,QgsCoordinateReferenceSystem('EPSG:4326'),QgsProject.instance()).transform(rect.center())
                    zone=min(60,max(1,int((center.x()+180)/6)+1))
                    default_crs='EPSG:'+str((32600 if center.y()>=0 else 32700)+zone)
        except (AttributeError,RuntimeError):pass
        self.addParameter(QgsProcessingParameterVectorLayer('ROADS','QuickOSM Urban road layer',[QgsProcessing.TypeVectorLine],default_layer))
        self.addParameter(QgsProcessingParameterVectorLayer('BUILDINGS','QuickOSM building polygon layer (optional, exported whole)',[QgsProcessing.TypeVectorPolygon],optional=True))
        self.addParameter(QgsProcessingParameterField('HEIGHT_FIELD','Building height column (optional; else OSM height, then building:levels x 3 m, else 6-8 m)',parentLayerParameterName='BUILDINGS',type=QgsProcessingParameterField.Any,optional=True))
        self.addParameter(QgsProcessingParameterEnum('SCOPE','Export area',['Full road layer (ignore screen zoom)','Specified extent / canvas'],defaultValue=0))
        self.addParameter(QgsProcessingParameterExtent('EXTENT','Extent (only for Specified extent)',optional=True))
        self.addParameter(QgsProcessingParameterCrs('TARGET_CRS','Export CRS (local projected metres)',defaultValue=default_crs))
        self.addParameter(QgsProcessingParameterEnum('WIDTH_MODE','Road widths',['Match QGIS style at reference scale (visual context)','Physical / estimated widths (OSM, lanes, road class)'],defaultValue=0))
        self.addParameter(QgsProcessingParameterNumber('SCALE','Reference map scale denominator (style rules and visual widths)',QgsProcessingParameterNumber.Double,default_scale,minValue=1))
        self.addParameter(QgsProcessingParameterString('NAME','Context name',defaultValue='Urban context'))
        self.addParameter(QgsProcessingParameterFileDestination('OUTPUT','Urban context package','JSON files (*.json)'))
    def processAlgorithm(self,parameters,context,feedback):
        layer=self.parameterAsVectorLayer(parameters,'ROADS',context)
        extent=None;crs=None
        if self.parameterAsEnum(parameters,'SCOPE',context)==1:
            if not parameters.get('EXTENT'):raise QgsProcessingException('Specify an extent or choose Full road layer')
            crs=layer.crs();extent=self.parameterAsExtent(parameters,'EXTENT',context,crs)
        path=self.parameterAsFileOutput(parameters,'OUTPUT',context)
        width_mode='qgis_style' if self.parameterAsEnum(parameters,'WIDTH_MODE',context)==0 else 'physical'
        export_layer(layer,extent,crs,self.parameterAsDouble(parameters,'SCALE',context),path,feedback,self.parameterAsString(parameters,'NAME',context),self.parameterAsCrs(parameters,'TARGET_CRS',context),width_mode,self.parameterAsVectorLayer(parameters,'BUILDINGS',context),self.parameterAsString(parameters,'HEIGHT_FIELD',context))
        return {'OUTPUT':path}
