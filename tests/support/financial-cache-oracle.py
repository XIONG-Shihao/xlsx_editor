"""Independent exact typed cache/source package proof, not a vendor oracle.
The17 rational expectations are mathematical. Remaining5 caches are checked
against the declared private evaluator bits to prove serialization integrity.
"""
import copy,hashlib,json,math,struct,sys,xml.etree.ElementTree as E,zipfile
N='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
def tree(e):return(e.tag,sorted(e.attrib.items()),e.text or '',e.tail or '',tuple(tree(c) for c in e))
def cells(e):
 d=e.findall(N+'sheetData');assert len(d)==1
 out={}
 for row in d[0]:
  assert row.tag==N+'row'
  for c in row:
   assert c.tag==N+'c' and c.get('r') and c.get('r') not in out
   out[c.get('r')]=c
 return out
def mask(root,targets):
 root=copy.deepcopy(root);data=root.find(N+'sheetData')
 for row in list(data):
  for cell in list(row):
   if cell.get('r') in targets:row.remove(cell)
  if not list(row) and set(row.attrib)=={'r'} and not(row.text or '').strip():data.remove(row)
 return tree(root)
def bits(v):return struct.pack('>d',v).hex()
source,saved,expected=sys.argv[1:4];expected=json.load(open(expected));assert expected['schema']=='workbook.xlsx.financial-private-cache-expectation.v1'
targets={c['cell'] for c in expected['cases']};assert len(targets)==len(expected['cases'])
result=[]
with zipfile.ZipFile(source)as a,zipfile.ZipFile(saved)as b:
 assert len(a.namelist())==len(set(a.namelist())) and len(b.namelist())==len(set(b.namelist()))
 assert set(a.namelist())==set(b.namelist())
 for part in a.namelist():
  if part!='xl/worksheets/sheet1.xml':assert a.read(part)==b.read(part),part
 before=E.fromstring(a.read('xl/worksheets/sheet1.xml'));after=E.fromstring(b.read('xl/worksheets/sheet1.xml'));assert mask(before,targets)==mask(after,targets),'unrelated worksheet/source delta'
 indexed=cells(after)
 for case in expected['cases']:
  cell=indexed[case['cell']];f=cell.findall(N+'f');v=cell.findall(N+'v');assert len(f)==len(v)==1 and not f[0].attrib and f[0].text==case['formula'],case['cell']
  assert set(cell.attrib)<= {'r','s','t'} and len(cell)==2
  if case['kind']=='number':
   assert cell.get('t') in (None,'n');actual=float(v[0].text);assert math.isfinite(actual);assert bits(actual)==case['bits'],case['cell']
  else:assert cell.get('t')=='str' and v[0].text==case['value']
  result.append({'cell':case['cell'],'formula':f[0].text,'kind':case['kind'],'raw':v[0].text,'bits':bits(actual)if case['kind']=='number' else None,'expectationLayer':case['expectationLayer']})
print(json.dumps({'schema':'workbook.xlsx.financial-private-cache-readback.v1','passed':True,'excelRuntimeEvidence':False,'targets':result,'unrelatedPartsAndCellsExact':True,'sourceSha256':hashlib.sha256(open(source,'rb').read()).hexdigest(),'savedSha256':hashlib.sha256(open(saved,'rb').read()).hexdigest()}))
