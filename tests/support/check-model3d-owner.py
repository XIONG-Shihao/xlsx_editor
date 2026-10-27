"""Saved package/graph oracle, separate from JS/native admission and rendering."""
import sys,json,zipfile,hashlib,xml.etree.ElementTree as E,tempfile,pathlib,subprocess
from decimal import Decimal
original,saved,expected=sys.argv[1:];rows=json.load(open(expected));D='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';M='http://schemas.microsoft.com/office/drawing/2017/model3d';A='http://schemas.openxmlformats.org/drawingml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';P='http://schemas.openxmlformats.org/package/2006/relationships';T='http://schemas.openxmlformats.org/package/2006/content-types';S='http://schemas.openxmlformats.org/spreadsheetml/2006/main';REL='http://schemas.microsoft.com/office/2017/06/relationships/model3d'
with zipfile.ZipFile(original) as old,zipfile.ZipFile(saved) as new:
 assert len(new.namelist())==len(set(new.namelist()))
 for part in old.namelist():
  if part.startswith('custom/'):assert new.read(part)==old.read(part)
 def cells(z,part):
  tree=E.fromstring(z.read(part))
  for c in tree.findall('.//{'+S+'}c'):
   if c.attrib.get('t','n')=='n':
    v=c.find('{'+S+'}v')
    if v is not None and v.text is not None:v.text=format(Decimal(v.text).normalize(),'f')
  return {c.attrib['r']:E.tostring(c) for c in tree.findall('.//{'+S+'}c')}
 for part in old.namelist():
  if part.startswith('xl/worksheets/') and part.endswith('.xml') and '/_rels/' not in part:assert cells(old,part)==cells(new,part),part
 models=[];assets=[];types=E.fromstring(new.read('[Content_Types].xml'));expected_by_id={r['id']:r for r in rows}
 with tempfile.TemporaryDirectory(prefix='saved-glb-oracle-') as temp:
  for part in new.namelist():
   if not(part.startswith('xl/drawings/') and part.endswith('.xml') and '/_rels/' not in part):continue
   canvas=E.fromstring(new.read(part));edges={e.attrib['Id']:e.attrib for e in E.fromstring(new.read(part.rsplit('/',1)[0]+'/_rels/'+part.rsplit('/',1)[1]+'.rels'))}
   for anchor in canvas:
    model=anchor.find('.//{'+M+'}model3d')
    if model is None:continue
    prop=anchor.find('.//{'+D+'}cNvPr');namespace,id=prop.attrib['name'].split(':',1);assert namespace == 'workbook';row=expected_by_id.pop(id);models.append(id)
    edge=edges[model.attrib['{'+R+'}embed']];assert edge['Type']==REL and 'TargetMode' not in edge
    asset='xl/media/'+edge['Target'].rsplit('/',1)[1];data=new.read(asset);digest=hashlib.sha256(data).hexdigest();assert asset=='xl/media/'+namespace+'-model3d-'+id+'-'+digest+'.glb';assets.append(asset)
    ct=[e for e in types if e.attrib.get('PartName')=='/'+asset];assert len(ct)==1 and ct[0].attrib['ContentType']=='model/gltf-binary'
    marker=model.find('.//{urn:'+namespace+':model3d:samples:v1}sample');wanted=row['model'];assert marker.attrib==dict(sample=wanted['sample'],color=wanted['color'],**{k:str(v) for k,v in wanted['rotation'].items()})
    assert [c.tag for c in model]==['{'+M+'}'+k for k in ['spPr','camera','trans','extLst','winViewport']]
    camera=model.find('{'+M+'}camera');assert camera.find('{'+M+'}pos').attrib==dict(x='0',y='0',z='180000000');assert camera.find('{'+M+'}orthographic/{'+M+'}sz').attrib==dict(n=str(min(row['anchor']['widthEmu'],row['anchor']['heightEmu'])),d=str(5*row['anchor']['heightEmu']))
    assert model.find('{'+M+'}trans/{'+M+'}meterPerModelUnit').attrib==dict(n='1',d='1')
    ext=anchor.find('{'+D+'}ext');assert (int(ext.attrib['cx']),int(ext.attrib['cy']))==(row['anchor']['widthEmu'],row['anchor']['heightEmu']);assert prop.attrib['descr']==row['alt']
    f=anchor.find('{'+D+'}from');col=int(f.find('{'+D+'}col').text);rownum=int(f.find('{'+D+'}row').text)+1;letters='';n=col+1
    while n:n,k=divmod(n-1,26);letters=chr(65+k)+letters
    assert letters+str(rownum)==row['anchor']['cell'];assert int(f.find('{'+D+'}colOff').text)==row['anchor']['dxEmu'];assert int(f.find('{'+D+'}rowOff').text)==row['anchor']['dyEmu']
    file=pathlib.Path(temp)/(id+'.glb');file.write_bytes(data);check=subprocess.run([sys.executable,str(pathlib.Path(__file__).with_name('check-model3d-glb.py')),str(file)],capture_output=True,text=True);assert check.returncode==0,check.stderr
    import struct
    length=struct.unpack_from('<I',data,12)[0];assert next(iter(json.loads(data[20:20+length])['extras'].values()))==wanted
 assert not expected_by_id
 assert set(assets)=={p for p in new.namelist() if p.endswith('.glb')}
 if not rows:assert not any(p.startswith('xl/drawings/') for p in new.namelist());assert not E.fromstring(new.read('xl/worksheets/sheet1.xml')).findall('{'+S+'}drawing')
print(json.dumps({'passed':True,'objects':len(models),'independentGLBChecks':len(assets),'sourceCellsOpaquePartIntact':True}))
