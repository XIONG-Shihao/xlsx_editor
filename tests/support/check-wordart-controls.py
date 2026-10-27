"""Independent WordArt text/run/effect/anchor/source graph oracle."""
import sys,json,zipfile,xml.etree.ElementTree as E,posixpath
D='{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}';S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';R='{http://schemas.openxmlformats.org/package/2006/relationships}';OR='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
profiles={'wordart-blue':('4472C4',None,False),'wordart-outline':(None,'C65911',False),'wordart-shadow':('7030A0',None,True)}
expected=json.load(open(sys.argv[3]));assert len({x['id'] for x in expected})==len(expected)
with zipfile.ZipFile(sys.argv[1]) as old,zipfile.ZipFile(sys.argv[2]) as z:
 files=[k for k in z.namelist() if k.startswith('xl/drawings/') and k.endswith('.xml')];assert len(files)==(1 if expected else 0)
 seen={};nonvisual=set()
 for file in files:
  draw=E.fromstring(z.read(file));assert draw.tag==D+'wsDr'
  for anchor in list(draw):
   assert anchor.tag==D+'oneCellAnchor';assert [n.tag for n in anchor]==[D+'from',D+'ext',D+'sp',D+'clientData']
   sp=anchor.find(D+'sp');p=sp.find(D+'nvSpPr/'+D+'cNvPr');id=p.get('name');assert id.startswith(('workbook:',));id=id.split(':',1)[1];assert id not in seen;number=p.get('id');assert number not in nonvisual;nonvisual.add(number)
   exp=next(x for x in expected if x['id']==id);color,outline,shadow=profiles[exp['sample']];assert p.get('descr')==exp['altText'];assert sp.find(D+'nvSpPr/'+D+'cNvSpPr').get('txBox')=='1'
   props=sp.find(D+'spPr');assert props.find(A+'prstGeom').get('prst')=='rect';assert props.find(A+'noFill') is not None;assert props.find(A+'ln/'+A+'noFill') is not None
   text=sp.find(D+'txBody');assert text.find(A+'bodyPr').attrib=={'fromWordArt':'1','wrap':'square','anchor':'ctr','lIns':'91440','rIns':'91440','tIns':'45720','bIns':'45720'};assert '\n'.join(p.find(A+'r/'+A+'t').text or '' for p in text.findall(A+'p'))==exp['text'];assert len(text.findall(A+'p'))<=4
   for paragraph in text.findall(A+'p'):
    assert paragraph.find(A+'pPr').get('algn')=='ctr';rpr=paragraph.find(A+'r/'+A+'rPr');assert rpr.attrib=={'lang':'en-US','sz':'3200','b':'1','i':'0'};assert rpr.find(A+'latin').get('typeface')=='Arial';fill=rpr.find(A+'solidFill/'+A+'srgbClr');assert (fill.get('val') if fill is not None else None)==color
    if color is None:assert rpr.find(A+'noFill') is not None
    line=rpr.find(A+'ln');assert line.get('w')=='12700';fill=line.find(A+'solidFill/'+A+'srgbClr');assert (fill.get('val') if fill is not None else None)==outline
    if outline is None:assert line.find(A+'noFill') is not None
    effect=rpr.find(A+'effectLst');assert effect is not None;assert len(effect)==(1 if shadow else 0)
    if shadow:
     sh=effect.find(A+'outerShdw');assert sh.attrib=={'blurRad':'19050','dist':'53882','dir':'2700000','sx':'100000','sy':'100000','kx':'0','ky':'0','algn':'ctr','rotWithShape':'0'};assert sh.find(A+'srgbClr').get('val')=='000000';assert sh.find(A+'srgbClr/'+A+'alpha').get('val')=='35000'
   a=exp['anchor'];letters=''.join(c for c in a['cell'] if c.isalpha());col=0
   for c in letters:col=col*26+ord(c)-64
   row=int(''.join(c for c in a['cell'] if c.isdigit()))-1
   assert {c.tag[len(D):]:int(c.text) for c in anchor.find(D+'from')}=={'col':col-1,'colOff':a['dxEmu'],'row':row,'rowOff':a['dyEmu']};assert anchor.find(D+'ext').attrib=={'cx':str(a['widthEmu']),'cy':str(a['heightEmu'])};assert props.find(A+'xfrm/'+A+'ext').attrib=={'cx':str(a['widthEmu']),'cy':str(a['heightEmu'])};seen[id]=True
 assert set(seen)=={x['id'] for x in expected};assert not any(k.startswith('xl/media/') for k in z.namelist())
 old_sheet=E.fromstring(old.read('xl/worksheets/sheet1.xml'));new_sheet=E.fromstring(z.read('xl/worksheets/sheet1.xml'));drawing=new_sheet.find(S+'drawing')
 if expected:
  assert drawing is not None;rels=E.fromstring(z.read('xl/worksheets/_rels/sheet1.xml.rels'));edge=next(n for n in rels if n.get('Id')==drawing.get(OR+'id'));assert edge.get('Type').endswith('/drawing');assert posixpath.normpath(posixpath.join('xl/worksheets',edge.get('Target')))==files[0];new_sheet.remove(drawing)
 else:assert drawing is None
 assert E.tostring(old_sheet)==E.tostring(new_sheet)
 for part in old.namelist():
  if part not in ['[Content_Types].xml','xl/worksheets/sheet1.xml','xl/workbook.xml','xl/calcChain.xml']:assert old.read(part)==z.read(part),part
 types=E.fromstring(z.read('[Content_Types].xml'));drawtypes=[n for n in types if n.get('ContentType','').endswith('drawing+xml')];assert len(drawtypes)==len(files)
 print(json.dumps({'passed':True,'samples':len(expected),'editableWordArtRuns':True,'worksheetAndOpaquePreserved':True,'graphBound':True,'noRasterSubstitute':True}))
