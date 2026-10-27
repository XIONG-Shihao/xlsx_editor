"""Independent real icon custom geometry/color/source package oracle."""
import sys,json,zipfile,xml.etree.ElementTree as E
D='{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}';S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
points={'icon-check':[(12,50),(35,73),(88,20),(96,28),(35,89),(4,58)],'icon-star':[(50,4),(61,36),(96,36),(68,57),(79,91),(50,70),(21,91),(32,57),(4,36),(39,36)],'icon-plus':[(40,5),(60,5),(60,40),(95,40),(95,60),(60,60),(60,95),(40,95),(40,60),(5,60),(5,40),(40,40)]}
with zipfile.ZipFile(sys.argv[1]) as old,zipfile.ZipFile(sys.argv[2]) as z:
 files=[n for n in z.namelist() if n.startswith('xl/drawings/') and n.endswith('.xml')];assert len(files)==1;draw=E.fromstring(z.read(files[0]));assert len(draw)==1;anchor=draw[0];assert anchor.tag==D+'oneCellAnchor';assert {n.tag[len(D):]:int(n.text) for n in anchor.find(D+'from')}=={'col':1,'colOff':100,'row':3,'rowOff':200};assert anchor.find(D+'ext').attrib=={'cx':'1143000','cy':'571500'}
 sp=anchor.find(D+'sp');assert sp.find(D+'txBody') is None;assert sp.find(D+'nvSpPr/'+D+'cNvPr').get('descr')=='Edited icon alternative text';assert sp.find(D+'nvSpPr/'+D+'cNvSpPr').get('txBox')=='0';props=sp.find(D+'spPr');assert props.find(A+'prstGeom') is None;assert props.find(A+'solidFill/'+A+'srgbClr').get('val')==sys.argv[4];assert props.find(A+'ln/'+A+'noFill') is not None
 geom=props.find(A+'custGeom');assert geom.find(A+'rect').attrib=={'l':'0','t':'0','r':'100','b':'100'};path=geom.find(A+'pathLst/'+A+'path');assert path.attrib=={'w':'100','h':'100','stroke':'0','extrusionOk':'0'};assert path[-1].tag==A+'close';actual=[]
 for i,node in enumerate(list(path)[:-1]):
  assert node.tag==A+('moveTo' if i==0 else 'lnTo');pt=node.find(A+'pt');actual.append((int(pt.get('x')),int(pt.get('y'))))
 assert actual==points[sys.argv[3]];old_sheet=E.fromstring(old.read('xl/worksheets/sheet1.xml'));new_sheet=E.fromstring(z.read('xl/worksheets/sheet1.xml'));new_sheet.remove(new_sheet.find(S+'drawing'));assert E.tostring(old_sheet)==E.tostring(new_sheet)
 for part in old.namelist():
  if part not in ['[Content_Types].xml','xl/worksheets/sheet1.xml','xl/workbook.xml','xl/calcChain.xml']:assert old.read(part)==z.read(part),part
 assert not any(n.startswith('xl/media/') for n in z.namelist());print(json.dumps({'passed':True,'sample':sys.argv[3],'color':sys.argv[4],'realCustomVector':True,'noTextOrRasterSurrogate':True,'sourceWorksheetPreserved':True}))
