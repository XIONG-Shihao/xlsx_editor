"""Independent SpreadsheetML/VML relationships, text/location and worksheet facts."""
import sys,json,zipfile,xml.etree.ElementTree as E,posixpath
S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}';P='{http://schemas.openxmlformats.org/package/2006/relationships}';V='{urn:schemas-microsoft-com:vml}';X='{urn:schemas-microsoft-com:office:excel}';T='{http://schemas.openxmlformats.org/package/2006/content-types}'
expected=json.load(open(sys.argv[3]));sheet='xl/worksheets/sheet1.xml';relpart='xl/worksheets/_rels/sheet1.xml.rels'
with zipfile.ZipFile(sys.argv[1]) as original,zipfile.ZipFile(sys.argv[2]) as saved:
 doc=E.fromstring(saved.read(sheet));old=E.fromstring(original.read(sheet));rels=E.fromstring(saved.read(relpart)) if relpart in saved.namelist() else E.Element('Relationships');links={r.attrib['Type'].rsplit('/',1)[-1]:r for r in rels};ld=doc.find(S+'legacyDrawing')
 if expected:
  assert ld is not None and 'comments' in links and 'vmlDrawing' in links;assert ld.attrib[R+'id']==links['vmlDrawing'].attrib['Id']
  def target(r):
   t=r.attrib['Target'];return t[1:] if t.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(sheet),t))
  c=target(links['comments']);v=target(links['vmlDrawing']);comments=E.fromstring(saved.read(c));authors=[a.text or '' for a in comments.find(S+'authors')];notes=comments.find(S+'commentList');shapes=E.fromstring(saved.read(v)).findall(V+'shape');assert len(notes)==len(shapes)==len(expected)
  for i,(n,shape,want) in enumerate(zip(notes,shapes,expected)):
   assert n.attrib['ref']==want['cell'];assert authors[int(n.attrib['authorId'])]==want['author'];assert n.find(S+'text/'+S+'t').text==want['text'];assert 'shapeId' not in n.attrib;assert shape.attrib['id']=='_x0000_s'+str(1025+i);assert 'visibility:hidden' in shape.attrib['style'];client=shape.find(X+'ClientData');assert client.attrib=={'ObjectType':'Note'}
   letters=''.join(x for x in want['cell'] if x.isalpha());number=int(''.join(x for x in want['cell'] if x.isdigit()));col=0
   for x in letters:col=col*26+ord(x)-64
   assert int(client.find(X+'Row').text)==number-1;assert int(client.find(X+'Column').text)==col-1;anchor=[int(x) for x in client.find(X+'Anchor').text.split(',')];assert len(anchor)==8 and 0<=anchor[0]<anchor[4]<16384 and 0<=anchor[2]<anchor[6]<1048576
  types=E.fromstring(saved.read('[Content_Types].xml'));ct={x.attrib.get('PartName'):x.attrib.get('ContentType')for x in types};assert ct['/'+c]=='application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';assert ct['/'+v]=='application/vnd.openxmlformats-officedocument.vmlDrawing'
 else:
  assert ld is None and 'comments' not in links and 'vmlDrawing' not in links;assert not any('/workbook-notes-' in x for x in saved.namelist())
 oldcells={c.attrib['r']:c for c in old.findall(S+'sheetData/'+S+'row/'+S+'c')};newcells={c.attrib['r']:c for c in doc.findall(S+'sheetData/'+S+'row/'+S+'c')};assert oldcells.keys()==newcells.keys()
 for ref,c in oldcells.items():
  n=newcells[ref];assert c.attrib==n.attrib
  if c.find(S+'f') is not None:assert c.find(S+'f').text==n.find(S+'f').text;assert float(n.find(S+'v').text)==6.0
  else:assert E.tostring(c)==E.tostring(n)
 for name in original.namelist():
  if name in [sheet,relpart,'[Content_Types].xml','xl/workbook.xml']:continue
  assert original.read(name)==saved.read(name),name
print(json.dumps({'passed':True,'notes':len(expected),'realCommentsAndVml':True,'cellsAndOpaquePartsPreserved':True}))
