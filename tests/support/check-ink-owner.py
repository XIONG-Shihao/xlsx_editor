"""Independent saved owner DrawingML paths and worksheet facts; no production parser."""
import sys,json,zipfile,xml.etree.ElementTree as E,posixpath
D='{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}';S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}';P='{http://schemas.openxmlformats.org/package/2006/relationships}';T='{http://schemas.openxmlformats.org/package/2006/content-types}'
styles={'black':('000000',19050,100000,'rnd'),'red':('D32F2F',19050,100000,'rnd'),'yellow-highlighter':('FFFF00',114300,30000,'flat')}
expected=json.load(open(sys.argv[3]));wanted={x['id']:x for x in expected['strokes']};seen=set();records=[]
with zipfile.ZipFile(sys.argv[1]) as old,zipfile.ZipFile(sys.argv[2]) as saved:
 for name in saved.namelist():
  if not name.startswith('xl/drawings/') or not name.endswith('.xml') or '/_rels/' in name:continue
  root=E.fromstring(saved.read(name))
  if root.tag!=D+'wsDr':continue
  for node in root:
   props=node.find(D+'sp/'+D+'nvSpPr/'+D+'cNvPr')
   if props is None:continue
   identity=props.get('name','').removeprefix('workbook:')
   if identity not in wanted:continue
   assert identity not in seen;seen.add(identity);w=wanted[identity];model=w['stroke'];anchor=w['anchor'];color,width,alpha,cap=styles[model['pen']]
   assert props.get('descr')==w['altText'];assert node.tag==D+'oneCellAnchor';letters=''.join(c for c in anchor['cell'] if c.isalpha());row=int(''.join(c for c in anchor['cell'] if c.isdigit()))-1;col=0
   for c in letters:col=col*26+ord(c)-64
   from_=node.find(D+'from');assert {x.tag.removeprefix(D):int(x.text)for x in from_}=={'col':col-1,'row':row,'colOff':anchor['dxEmu'],'rowOff':anchor['dyEmu']};assert node.find(D+'ext').attrib=={'cx':str(model['widthEmu']),'cy':str(model['heightEmu'])}
   pr=node.find(D+'sp/'+D+'spPr');path=pr.find(A+'custGeom/'+A+'pathLst/'+A+'path');assert path.attrib=={'w':str(model['widthEmu']),'h':str(model['heightEmu']),'fill':'none','stroke':'1','extrusionOk':'0'};assert [x.tag for x in path]==[A+'moveTo']+[A+'lnTo']*(len(model['points'])-1);assert [[int(x[0].get('x')),int(x[0].get('y'))]for x in path]==model['points'];assert path.find(A+'close') is None
   line=pr.find(A+'ln');assert line.attrib=={'w':str(width),'cap':cap,'cmpd':'sng','algn':'ctr'};rgb=line.find(A+'solidFill/'+A+'srgbClr');assert rgb.get('val')==color and rgb.find(A+'alpha').get('val')==str(alpha);assert pr.find(A+'noFill') is not None and node.find(D+'pic') is None
   records.append({'id':identity,'pen':model['pen'],'points':len(model['points']),'anchor':anchor['cell'],'openVector':True})
  for sheet in [s for s in saved.namelist() if s.startswith('xl/worksheets/') and s.endswith('.xml') and '/_rels/' not in s]:
   doc=E.fromstring(saved.read(sheet));drawing=doc.find(S+'drawing')
   if drawing is None:continue
   relpart=posixpath.join(posixpath.dirname(sheet),'_rels',posixpath.basename(sheet)+'.rels');rels=E.fromstring(saved.read(relpart));edges=[x for x in rels if x.get('Id')==drawing.get(R+'id')];assert len(edges)==1;edge=edges[0];assert edge.get('Type')==R[1:-1]+'/drawing' and edge.get('TargetMode') is None
   target=edge.get('Target');resolved=target[1:] if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(sheet),target));assert resolved in saved.namelist()
 assert seen==set(wanted),(seen,wanted)
 for sheet in [s for s in old.namelist() if s.startswith('xl/worksheets/') and s.endswith('.xml') and '/_rels/' not in s]:
  source=E.fromstring(old.read(sheet));dest=E.fromstring(saved.read(sheet));before={c.get('r'):c for c in source.findall(S+'sheetData/'+S+'row/'+S+'c')};after={c.get('r'):c for c in dest.findall(S+'sheetData/'+S+'row/'+S+'c')};assert before.keys()==after.keys()
  for ref,c in before.items():
   n=after[ref];assert c.attrib==n.attrib
   if c.find(S+'f') is not None:
    assert E.tostring(c.find(S+'f'))==E.tostring(n.find(S+'f'))
    if ref in expected.get('formulaCaches',{}):assert float(n.find(S+'v').text)==expected['formulaCaches'][ref]
   else:assert E.tostring(c)==E.tostring(n)
 for name in old.namelist():
  if name in ['[Content_Types].xml','xl/workbook.xml'] or name.startswith('xl/worksheets/') or name.startswith('xl/drawings/'):continue
  assert saved.read(name)==old.read(name),name
 for name in expected.get('preservedParts',[]):assert saved.read(name)==old.read(name),name
 if not expected['strokes'] and expected.get('noDrawings'):assert not any(n.startswith('xl/drawings/')for n in saved.namelist())
print(json.dumps({'passed':True,'strokes':records,'worksheetFactsAndOpaquePreserved':True,'realVectorPaths':True,'scope':'Actual native/Document saved paths and source facts; no pointer/editor/Excel interchange claim'}))
