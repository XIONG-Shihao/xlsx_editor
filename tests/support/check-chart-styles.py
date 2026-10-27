"""Independent actual DrawingML palette/fill/line and refreshed-cache check."""
import sys,json,zipfile,xml.etree.ElementTree as E
C='{http://schemas.openxmlformats.org/drawingml/2006/chart}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
colorful=['2563EB','E86125','16A34A','9333EA','CA8A04','0891B2','DC2626','64748B'];mono=['217346','42875F','66A27F','85B998','175431','327149','50946A','749F84'];rows=json.load(open(sys.argv[3]))
with zipfile.ZipFile(sys.argv[1]) as old,zipfile.ZipFile(sys.argv[2]) as saved:
 for name in old.namelist():
  if name.startswith('custom/'):assert old.read(name)==saved.read(name),name
 for row in rows:
  tree=E.fromstring(saved.read('xl/charts/workbook-chart-'+row['id']+'.xml'));assert tree.tag==C+'chartSpace';kind=row['kind'];plot=tree.find(C+'chart/'+C+'plotArea');chart=plot.find(C+('lineChart' if kind=='line' else 'barChart'));assert chart is not None
  if kind!='line':assert chart.find(C+'barDir').attrib['val']==('bar' if kind=='bar' else 'col')
  series=chart.findall(C+'ser');assert len(series)==len(row['values']);palette=mono if row['style']=='monochrome' else colorful;outline=row['style']=='outline'
  for i,s in enumerate(series):
   assert s.find(C+'idx').attrib=={'val':str(i)};props=s.find(C+'spPr');assert props is not None and props.attrib=={};assert [n.tag for n in props]==[A+('noFill' if outline else 'solidFill'),A+'ln']
   if not outline:assert props.find(A+'solidFill/'+A+'srgbClr').attrib=={'val':palette[i]}
   line=props.find(A+'ln');assert line.attrib=={'w':'19050'};assert [n.tag for n in line]==[A+'solidFill',A+'prstDash'];assert line.find(A+'solidFill/'+A+'srgbClr').attrib=={'val':palette[i]};assert line.find(A+'prstDash').attrib=={'val':'dash' if outline else 'solid'}
   cache=s.find(C+'val/'+C+'numRef/'+C+'numCache');assert int(cache.find(C+'ptCount').attrib['val'])==len(row['values'][i]);points={int(p.attrib['idx']):float(p.find(C+'v').text) for p in cache.findall(C+'pt')};assert [points.get(j) for j in range(len(row['values'][i]))]==row['values'][i]
  assert len(tree.findall('.//'+C+'spPr'))==len(series)
print(json.dumps({'passed':True,'charts':len(rows),'realSeriesProperties':True,'cacheValuesMatch':True,'opaqueMembersPreserved':True}))
