"""Versioned, Revit-free Urban context package validation and placement."""
import io
import json
import math
import os
import tempfile
import uuid
from collections import deque

try:
    STRING_TYPES=(basestring,)
except NameError:
    STRING_TYPES=(str,)

FORMAT='WWPTools.UrbanContext'
VERSION=1
MAX_FILE_BYTES=64*1024*1024
MAX_VERTICES=250000

def _text(value,label,limit=200):
    if not isinstance(value,STRING_TYPES) or not value.strip() or len(value)>limit or '\x00' in value:
        raise ValueError('Invalid '+label)
    return value

def _number(value,label):
    if isinstance(value,bool) or not isinstance(value,(int,float)) or math.isnan(value) or math.isinf(value):
        raise ValueError('Invalid finite number: '+label)
    return value

def _point(point,label,local=False):
    if not isinstance(point,(list,tuple)) or len(point)!=2:
        raise ValueError('Invalid '+label)
    for n in point:_number(n,label)
    if local and math.hypot(*point)>10000:
        raise ValueError('Context extends beyond 10 km from its local origin; export a smaller area')

def _rgb(values):
    if not isinstance(values,list) or len(values)!=3 or any(isinstance(v,bool) or not isinstance(v,int) or not 0<=v<=255 for v in values):
        raise ValueError('Invalid RGB colour')

def validate_package(data):
    """Raise ValueError before edits if the package is unsupported or unsafe to import."""
    if not isinstance(data,dict) or data.get('format')!=FORMAT or data.get('version')!=VERSION:
        raise ValueError('Unsupported Urban context package format/version')
    _text(data.get('dataset_id'),'dataset identity')
    _text(data.get('name'),'dataset name')
    frame=data.get('frame') or {}
    if frame.get('units')!='metres':raise ValueError('Package coordinates must use metres')
    crs=_text(frame.get('crs'),'projected CRS',2000)
    if crs in ('EPSG:4326','EPSG:3857'):raise ValueError('Use a local projected metric CRS, not geographic or Web Mercator')
    _point(frame.get('origin'),'projected origin')
    cats=data.get('categories')
    if not isinstance(cats,list) or not 1<=len(cats)<=256:raise ValueError('Invalid category count')
    ids=set()
    for c in cats:
        cid=_text(c.get('id'),'category identity')
        if cid in ids:raise ValueError('Duplicate category identity')
        ids.add(cid);_text(c.get('name'),'category name');_rgb(c.get('rgb'));_rgb(c.get('outline_rgb',c.get('rgb')))
    surfaces=data.get('surfaces')
    if not isinstance(surfaces,list) or not 1<=len(surfaces)<=10000:raise ValueError('Invalid/empty surface count')
    sids=set();vertices=0
    for s in surfaces:
        sid=_text(s.get('id'),'surface identity')
        if sid in sids:raise ValueError('Duplicate surface identity')
        sids.add(sid)
        if s.get('category') not in ids:raise ValueError('Surface references a missing category')
        rings=s.get('rings')
        if not isinstance(rings,list) or not rings:raise ValueError('Surface has no polygon rings')
        for ring in rings:
            if not isinstance(ring,list) or len(ring)<4 or ring[0]!=ring[-1]:raise ValueError('Polygon rings must be closed with at least three vertices')
            vertices+=len(ring)
            if vertices>MAX_VERTICES:raise ValueError('Too many polygon vertices; export a smaller area')
            for point in ring:_point(point,'polygon vertex',True)
            area=sum(ring[i][0]*ring[i+1][1]-ring[i+1][0]*ring[i][1] for i in range(len(ring)-1))
            if abs(area)<1e-8:raise ValueError('Degenerate polygon ring')
    labels=data.get('labels',[])
    if not isinstance(labels,list) or len(labels)>5000:raise ValueError('Invalid label count')
    lids=set()
    for label in labels:
        lid=_text(label.get('id'),'label identity')
        if lid in lids:raise ValueError('Duplicate label identity')
        lids.add(lid);_text(label.get('text'),'label text',2000)
        _point(label.get('point'),'label point',True)
        _number(label.get('rotation_rad',0),'label rotation')
    return data

def local_to_revit(point,frame,placement):
    """Map package metres to internal feet using the first import's fixed reference."""
    if placement['crs']!=frame['crs']:raise ValueError('The package CRS changed; use a new export file or retain the original CRS')
    x=point[0]+frame['origin'][0]-placement['source_origin'][0]
    y=point[1]+frame['origin'][1]-placement['source_origin'][1]
    angle=placement.get('rotation_rad',0)
    c,s=math.cos(angle),math.sin(angle)
    anchor=placement['anchor_ft']
    result=[anchor[0]+(x*c-y*s)/0.3048,anchor[1]+(x*s+y*c)/0.3048,anchor[2]]
    if math.sqrt(sum(n*n for n in result))>16000/0.3048:raise ValueError('Placed context is too far from Revit internal origin')
    return result

def clean_ring(ring,tolerance):
    """Remove sub-tolerance edges and returning spikes without mutating a closed ring."""
    _number(tolerance,'ring tolerance')
    if tolerance<=0:raise ValueError('Ring tolerance must be positive')
    limit=tolerance*tolerance
    def close(a,b):return (a[0]-b[0])**2+(a[1]-b[1])**2<=limit
    vertices=[]
    for point in ring[:-1]:
        if vertices and close(vertices[-1],point):continue
        if len(vertices)>1 and close(vertices[-2],point):
            vertices.pop()
            continue
        vertices.append(list(point))
    vertices=deque(vertices)
    while len(vertices)>2:
        if close(vertices[-1],vertices[0]):vertices.pop()
        elif close(vertices[-2],vertices[0]):
            vertices.pop();vertices.pop()
        elif close(vertices[-1],vertices[1]):
            vertices.pop();vertices.popleft()
        else:break
    if len(vertices)<3:raise ValueError('Polygon ring collapsed below Revit geometry tolerance')
    return list(vertices)+[list(vertices[0])]

def filter_categories(package,category_ids):
    """Return a validated group selection without mutating the source package."""
    validate_package(package)
    selected=set(category_ids)
    known=set(c['id'] for c in package['categories'])
    if not selected or not selected.issubset(known):raise ValueError('Select at least one available road group')
    data=json.loads(json.dumps(package,ensure_ascii=False))
    data['categories']=[c for c in data['categories'] if c['id'] in selected]
    data['surfaces']=[s for s in data['surfaces'] if s['category'] in selected]
    associated={}
    for source in package.get('sources',[]):
        associated.setdefault(source.get('name',''),set()).add(source.get('category'))
    labels=[]
    for label in data.get('labels',[]):
        groups=set(label.get('categories',[])) or associated.get(label['text'],set())
        if groups.intersection(selected) or (not groups and selected==known):labels.append(label)
    data['labels']=labels
    if 'sources' in data:data['sources']=[s for s in data['sources'] if s.get('category') in selected]
    data['selected_category_ids']=sorted(selected)
    return validate_package(data)

def load_package(path):
    if os.path.getsize(path)>MAX_FILE_BYTES:raise ValueError('Urban package exceeds 64 MB')
    with io.open(path,'r',encoding='utf-8') as stream:data=json.load(stream)
    return validate_package(data)

def write_package(path,data):
    """Validate and atomically replace a package; exports to the same file retain identity."""
    data=json.loads(json.dumps(data,ensure_ascii=False))
    if os.path.exists(path):
        previous=load_package(path)
        data['dataset_id']=previous['dataset_id']
    elif not data.get('dataset_id'):
        data['dataset_id']=str(uuid.uuid4())
    validate_package(data)
    content=json.dumps(data,ensure_ascii=False,allow_nan=False,indent=2)
    if len(content.encode('utf-8'))>MAX_FILE_BYTES:raise ValueError('Urban package exceeds 64 MB')
    folder=os.path.dirname(os.path.abspath(path))
    fd,temp=tempfile.mkstemp(prefix='.urban-',suffix='.tmp',dir=folder)
    try:
        with io.open(fd,'w',encoding='utf-8',closefd=True) as stream:stream.write(content)
        if hasattr(os,'replace'):os.replace(temp,path)
        else:
            import System
            if os.path.exists(path):System.IO.File.Replace(temp,path,None)
            else:System.IO.File.Move(temp,path)
    finally:
        if os.path.exists(temp):os.remove(temp)
    return data
