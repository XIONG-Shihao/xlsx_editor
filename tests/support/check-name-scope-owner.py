"""Independent caches and unrelated XML/ZIP, allowing inert default empty cells."""
import sys,json,zipfile,struct,re,xml.etree.ElementTree as E
source,saved,expected_path=sys.argv[1:]
expected=json.load(open(expected_path));N='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
def parts(file):
 with zipfile.ZipFile(file) as z:
  assert len(z.namelist())==len(set(z.namelist()));return {name:z.read(name) for name in z.namelist()}
def tree(node):return(node.tag,tuple(sorted(node.attrib.items())),node.text or '',node.tail or '',tuple(tree(c)for c in node))
a,b=parts(source),parts(saved);assert set(a)==set(b);observed=[];raw_exact=0;semantic_worksheets=[];dimension_normalizations=[]
for name in a:
 changes=expected.get(name,{})
 if not changes and a[name]==b[name]:raw_exact+=1;continue
 old,new=E.fromstring(a[name]),E.fromstring(b[name]);assert old.tag==new.tag==N+'worksheet','Unrelated non-worksheet part changed: '+name
 if not changes:semantic_worksheets.append(name)
 old_dimension,new_dimension=old.find(N+'dimension'),new.find(N+'dimension')
 if old_dimension is not None and new_dimension is not None and old_dimension.get('ref')!=new_dimension.get('ref'):
  # dimension is a cached extent hint. The established writer recomputes it;
  # accept only the independently derived bounds of actual saved cells.
  assert set(old_dimension.attrib)==set(new_dimension.attrib)=={'ref'} and len(old_dimension)==len(new_dimension)==0
  points=[]
  for c in new.findall(N+'sheetData/'+N+'row/'+N+'c'):
   if len(c)==0 and not (c.text or '').strip() and set(c.attrib).issubset({'r','s'}) and c.get('s','0')=='0':continue
   m=re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]*)',c.get('r',''));assert m
   col=0
   for ch in m[1]:col=col*26+ord(ch)-64
   points.append((col,int(m[2])))
  assert points,'Changed extent of an empty worksheet needs an explicit profile'
  def column(n):
   text=''
   while n:n,rest=divmod(n-1,26);text=chr(65+rest)+text
   return text
  first=column(min(c for c,r in points))+str(min(r for c,r in points));last=column(max(c for c,r in points))+str(max(r for c,r in points))
  allowed={first+':'+last}
  if first==last:allowed.add(first)
  assert new_dimension.get('ref') in allowed,'Saved extent does not match actual cells'
  dimension_normalizations.append({'part':name,'prior':old_dimension.get('ref'),'saved':new_dimension.get('ref')})
  old.remove(old_dimension);new.remove(new_dimension)
 old_cells={c.get('r'):c for c in old.findall(N+'sheetData/'+N+'row/'+N+'c')};new_cells={c.get('r'):c for c in new.findall(N+'sheetData/'+N+'row/'+N+'c')}
 for ref,want in changes.items():
  c=new_cells[ref];f=c.findall(N+'f');v=c.findall(N+'v');assert len(f)==len(v)==1 and f[0].text==want['formula'],ref
  assert c.get('t','n')=='n';assert struct.pack('>d',float(v[0].text))==struct.pack('>d',want['value']),ref
  if want.get('retained'):
   prior=old_cells[ref];old_v=prior.findall(N+'v');assert len(old_v)==1
   assert struct.pack('>d',float(old_v[0].text))==struct.pack('>d',float(v[0].text)),ref
   prior.remove(old_v[0]);c.remove(v[0]);assert tree(prior)==tree(c),'Retained cache non-value field changed: '+ref
  observed.append({'part':name,'reference':ref,'formula':f[0].text,'value':float(v[0].text),'bits':struct.pack('>d',float(v[0].text)).hex()})
 for node in [old,new]:
  for row in node.findall(N+'sheetData/'+N+'row'):
   for c in list(row):
    if c.tag==N+'c' and c.get('r') in changes:row.remove(c)
  # The established writer omits redundant empty cells in style 0. Permit
  # only those inert markers; real styles, cell metadata and typed content
  # remain exact. This is semantic normalization, not a general XML waiver.
  for row in node.findall(N+'sheetData/'+N+'row'):
   for c in list(row):
    if c.tag==N+'c' and len(c)==0 and not (c.text or '').strip() and set(c.attrib).issubset({'r','s'}) and c.get('s','0')=='0':row.remove(c)
  # Insertion or inert-cell omission can leave ordinary empty rows.
  data=node.find(N+'sheetData')
  for row in list(data):
   if row.tag==N+'row' and len(row)==0 and set(row.attrib)=={'r'}:data.remove(row)
 assert tree(old)==tree(new),'Unrelated worksheet field changed: '+name
print(json.dumps({'passed':True,'checkedCaches':observed,'unrelatedZipPartsExact':raw_exact,'unchangedWorksheetsSemanticallyVerified':semantic_worksheets,'normalizedWorksheetDimensions':dimension_normalizations,'workbookNamesAndRawScopesExact':a['xl/workbook.xml']==b['xl/workbook.xml']}))
