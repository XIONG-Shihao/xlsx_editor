"""Independent saved SmartArt semantics/placement/cache/ownership check; no production generator."""
from pathlib import Path
from decimal import Decimal
import sys,json,zipfile,xml.etree.ElementTree as E,posixpath,math
D='{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}'
G='{http://schemas.openxmlformats.org/drawingml/2006/diagram}'
S='{http://schemas.microsoft.com/office/drawing/2008/diagram}'
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
P='{http://schemas.openxmlformats.org/package/2006/relationships}'
T='{http://schemas.openxmlformats.org/package/2006/content-types}'
original,saved,expected=map(Path,sys.argv[1:])
with zipfile.ZipFile(original) as z:before={n:z.read(n) for n in z.namelist()}
with zipfile.ZipFile(saved) as z:parts={n:z.read(n) for n in z.namelist()}
trees={n:E.fromstring(b) for n,b in parts.items() if n.endswith('.xml') or n.endswith('.rels')}
rows=json.loads(expected.read_text());assert len(rows)==len({r['id'] for r in rows})
for name,b in before.items():
    if name.startswith('custom/'):assert parts[name]==b,name
    if name.startswith('xl/worksheets/sheet'):
        before_sheet=E.fromstring(b);after_sheet=trees[name]
        ns='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        for tree in (before_sheet,after_sheet):
            for node in list(tree):
                if node.tag==ns+'drawing':tree.remove(node)
        # Native calculation export may serialize an unchanged numeric cache as 43.0.
        # Compare its numeric value while retaining formula/text/style/row facts.
        for tree in (before_sheet,after_sheet):
            for cell in tree.findall('.//'+ns+'c'):
                if cell.get('t','n')=='n':
                    value=cell.find(ns+'v')
                    if value is not None and value.text is not None:value.text=format(Decimal(value.text).normalize(),'f')
        assert E.tostring(before_sheet)==E.tostring(after_sheet),name
incoming={}
for name,t in trees.items():
    if not name.endswith('.rels'):continue
    owner='' if name=='_rels/.rels' else name.replace('/_rels/','/')[:-5]
    for edge in t:
        if edge.get('TargetMode')=='External':continue
        target=edge.get('Target');to=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(owner),target))
        incoming[to]=incoming.get(to,0)+1
found={}
for name,t in trees.items():
    if not name.startswith('xl/drawings/') or not name.endswith('.xml') or '/_rels/' in name:continue
    for anchor in t:
        props=anchor.find(D+'graphicFrame/'+D+'nvGraphicFramePr/'+D+'cNvPr')
        if props is not None and props.get('name','').startswith(('workbook:',)):found[props.get('name').split(':',1)[1]]=(name,anchor,props)
assert set(found)=={r['id'] for r in rows}
for row in rows:
    id=row['id'];model=row['model'];kind=model['layout'];labels=[n['text'] for n in model['nodes']];n=len(labels)
    drawing,anchor,props=found[id];namespace=props.get('name').split(':',1)[0];placement=row['anchor'];assert anchor.tag==D+'oneCellAnchor';assert props.get('descr')==row['alt']
    letters=''.join(c for c in placement['cell'] if c.isalpha());col=0
    for c in letters:col=col*26+ord(c)-64
    first=anchor.find(D+'from');assert [int(first.find(D+k).text) for k in ['col','colOff','row','rowOff']]==[col-1,placement['dxEmu'],int(placement['cell'][len(letters):])-1,placement['dyEmu']]
    extent=anchor.find(D+'ext');assert [int(extent.get(k)) for k in ['cx','cy']]==[placement['widthEmu'],placement['heightEmu']]
    paths={role:f'xl/diagrams/{namespace}-smartart-{id}-{role}.xml' for role in ['data','layout','style','colors','drawing']}
    refs=anchor.find(D+'graphicFrame/'+A+'graphic/'+A+'graphicData/'+G+'relIds');rels={e.get('Id'):e for e in trees[drawing.replace('/'+drawing.split('/')[-1],'/_rels/'+drawing.split('/')[-1]+'.rels')]}
    for attr,role,suffix in [('dm','data','Data'),('lo','layout','Layout'),('qs','style','QuickStyle'),('cs','colors','Colors')]:
        edge=rels[refs.get(R+attr)];assert edge.get('Type').endswith('/diagram'+suffix);assert posixpath.normpath(posixpath.join(posixpath.dirname(drawing),edge.get('Target')))==paths[role];assert edge.get('TargetMode') is None
    for path in paths.values():assert incoming[path]==1
    data=trees[paths['data']];points=list(data.find(G+'ptLst'));ids={p.get('modelId') for p in points};assert len(ids)==len(points)
    nodes=[p for p in points if p.get('type')=='node'];assert [p.find(G+'t/'+A+'p/'+A+'r/'+A+'t').text for p in nodes]==labels
    for c in data.find(G+'cxnLst'):
        assert c.get('srcId') in ids and c.get('destId') in ids
        for key in ['parTransId','sibTransId']:
            if c.get(key):assert c.get(key) in ids
    assert trees[paths['layout']].get('uniqueId')==f'urn:{namespace}:smartart:layout:{kind}'
    assert trees[paths['layout']].find(G+'layoutNode/'+G+'alg').get('type')=={'process':'lin','hierarchy':'hierRoot','cycle':'cycle'}[kind]
    types=list(trees['[Content_Types].xml'])
    for role,path in paths.items():
        matches=[t for t in types if t.get('PartName')=='/'+path];assert len(matches)==1
        assert matches[0].get('ContentType')==('application/vnd.ms-office.drawingml.diagramDrawing+xml' if role=='drawing' else 'application/vnd.openxmlformats-officedocument.drawingml.diagram'+role.title()+'+xml')
    cache=trees[paths['drawing']].find(S+'spTree');shapes=cache.findall(S+'sp');text_shapes=[s for s in shapes if s.find(S+'txBody') is not None];lines=[s for s in shapes if s.find(S+'txBody') is None]
    assert [s.find(S+'txBody/'+A+'p/'+A+'r/'+A+'t').text for s in text_shapes]==labels
    assert len(lines)==(n if kind=='cycle' else n-1)
    presentation={p.get('modelId') for p in points if p.get('type')=='pres'};assert all(s.get('modelId') in presentation for s in shapes)
    w,h=placement['widthEmu'],placement['heightEmu'];boxes=[]
    for i in range(n):
        if kind=='process':
            step=920/n;bw=min(260,step-30);box=(40+i*step+(step-bw)/2,220,bw,160)
        elif kind=='hierarchy':
            step=920/(n-1);bw=min(260,step-24);box=(370 if i==0 else 40+(i-1)*step+(step-bw)/2,50 if i==0 else 380,260 if i==0 else bw,150)
        else:
            angle=-math.pi/2+2*math.pi*i/n;box=(500+math.cos(angle)*320-90,300+math.sin(angle)*180-55,180,110)
        boxes.append(box)
    for shape,box in zip(text_shapes,boxes):
        props=shape.find(S+'spPr');transform=props.find(A+'xfrm');offset=transform.find(A+'off');extent=transform.find(A+'ext')
        got=[int(offset.get('x')),int(offset.get('y')),int(extent.get('cx')),int(extent.get('cy'))]
        want=[math.floor(box[0]*w/1000+0.5),math.floor(box[1]*h/600+0.5),math.floor(box[2]*w/1000+0.5),math.floor(box[3]*h/600+0.5)]
        assert all(abs(a-b)<=1 for a,b in zip(got,want)),(got,want)
        assert props.find(A+'solidFill/'+A+'srgbClr').get('val')=='4472C4';run=shape.find(S+'txBody/'+A+'p/'+A+'r/'+A+'rPr');assert run.get('sz')=='1400' and run.find(A+'latin').get('typeface')=='Arial'
    data_rels=trees[f'xl/diagrams/_rels/{namespace}-smartart-{id}-data.xml.rels'];assert len(data_rels)==1 and data_rels[0].get('Type').endswith('/diagramDrawing') and data_rels[0].get('Target')==paths['drawing'].split('/')[-1]
if not rows:assert not any(name.startswith('xl/diagrams/') or name.startswith('xl/drawings/') for name in parts)
print(json.dumps({'passed':True,'objects':len(rows),'semanticGraphAndCache':True,'placementAndUnrelatedData':True}))
