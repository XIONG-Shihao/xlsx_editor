"""Independent editable shape geometry/text/package and original worksheet proof."""
import sys,json,zipfile,xml.etree.ElementTree as E
D='{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
expected={'text-plain':('rect',None,None,'1F2937','1'),'text-note':('rect','EAF2F8','4472C4','17365D','1'),'shape-rectangle':('rect','DCE6F1','4472C4','17365D','0'),'shape-ellipse':('ellipse','E2EFDA','70AD47','375623','0'),'shape-arrow':('rightArrow','FFF2CC','BF9000','7F6000','0')}
with zipfile.ZipFile(sys.argv[1]) as old,zipfile.ZipFile(sys.argv[2]) as z:
 files=[k for k in z.namelist() if k.startswith('xl/drawings/') and k.endswith('.xml')];assert len(files)==1
 draw=E.fromstring(z.read(files[0]));anchor=draw.find(D+'oneCellAnchor');assert len(list(draw))==1;assert anchor.find(D+'from/'+D+'col').text=='1' and anchor.find(D+'from/'+D+'row').text=='3';assert anchor.find(D+'from/'+D+'colOff').text=='100';assert anchor.find(D+'from/'+D+'rowOff').text=='200';assert anchor.find(D+'ext').attrib=={'cx':'1143000','cy':'571500'}
 sp=anchor.find(D+'sp');geom,bg,line,color,tx=expected[sys.argv[3]];assert sp is not None;assert sp.find(D+'nvSpPr/'+D+'cNvPr').get('descr')=='Edited alternative text';assert sp.find(D+'nvSpPr/'+D+'cNvSpPr').get('txBox')==tx;props=sp.find(D+'spPr');assert props.find(A+'prstGeom').get('prst')==geom
 fill=props.find(A+'solidFill/'+A+'srgbClr');assert (fill.get('val') if fill is not None else None)==bg;stroke=props.find(A+'ln/'+A+'solidFill/'+A+'srgbClr');assert (stroke.get('val') if stroke is not None else None)==line
 text=sp.find(D+'txBody');assert [p.find(A+'r/'+A+'t').text for p in text.findall(A+'p')]==['Edited "&é','Next'];assert all(p.find(A+'r/'+A+'rPr').get('sz')=='1400' for p in text.findall(A+'p'));assert all(p.find(A+'r/'+A+'rPr/'+A+'solidFill/'+A+'srgbClr').get('val')==color for p in text.findall(A+'p'))
 S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';old_sheet=E.fromstring(old.read('xl/worksheets/sheet1.xml'));new_sheet=E.fromstring(z.read('xl/worksheets/sheet1.xml'));new_sheet.remove(new_sheet.find(S+'drawing'));assert E.tostring(old_sheet)==E.tostring(new_sheet)
 for part in old.namelist():
  if part not in ['[Content_Types].xml','xl/worksheets/sheet1.xml','xl/workbook.xml','xl/calcChain.xml']:assert old.read(part)==z.read(part),part
 assert not any(k.startswith('xl/media/') for k in z.namelist());print(json.dumps({'passed':True,'sample':sys.argv[3],'editableDrawingML':True,'text':['Edited "&é','Next'],'worksheetAndOpaquePreserved':True,'noRasterSubstitute':True}))
