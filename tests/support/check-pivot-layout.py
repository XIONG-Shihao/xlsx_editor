"""Independent saved OOXML presentation/cached values and preservation oracle."""
import sys,json,zipfile,xml.etree.ElementTree as E
n='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
sample=sys.argv[3];grand=sys.argv[4]=='true'
compact,outline,style,stripes={'tabular-light':('0','0','PivotStyleLight16','0'),'compact-blue':('1','1','PivotStyleMedium9','1'),'outline-dark':('0','1','PivotStyleDark2','1')}[sample]
with zipfile.ZipFile(sys.argv[1]) as old,zipfile.ZipFile(sys.argv[2]) as z:
 namespace='workbook'
 t=E.fromstring(z.read('xl/pivotTables/'+namespace+'Pivot1.xml'));assert t.get('name')=='SalesPivot' and t.get('cacheId')=='1'
 for key,value in {'compact':compact,'compactData':compact,'outline':outline,'outlineData':outline,'rowGrandTotals':str(int(grand)),'colGrandTotals':str(int(grand))}.items():assert t.get(key)==value,key
 assert t.find(n+'location').get('ref')==('A1:B5' if grand else 'A1:B4')
 s=t.find(n+'pivotTableStyleInfo');assert s.get('name')==style and s.get('showRowStripes')==stripes
 f=t.find(n+'pivotFields')[0];assert f.get('compact')==compact and f.get('outline')==outline
 assert len(f.find(n+'items'))==3 and len(t.find(n+'rowItems'))==3+int(grand)
 assert t.find(n+'colItems').get('count')=='1' and len(t.find(n+'colItems')[0])==0
 for cache in ['xl/pivotCache/'+namespace+'CacheDefinition1.xml','xl/pivotCache/'+namespace+'CacheRecords1.xml']:assert old.read(cache)==z.read(cache)
 sheet=E.fromstring(z.read('xl/worksheets/sheet3.xml'));cells={c.get('r'):c for c in sheet.iter(n+'c')};read=lambda c:c.find(n+'is/'+n+'t').text if c.get('t')=='inlineStr' else int(c.find(n+'v').text)
 expected=[['Row Labels' if compact=='1' else 'Category','Sum of Amount'],['A',1],['B',5],['C',4]]+([['Grand Total',10]] if grand else [])
 actual=[[read(cells[col+str(i+1)]) for col in 'AB'] for i in range(len(expected))];assert actual==expected
 if not grand:assert 'A5' not in cells and 'B5' not in cells
 for part in old.namelist():
  if part not in ['xl/pivotTables/'+namespace+'Pivot1.xml','xl/worksheets/sheet3.xml','xl/workbook.xml','xl/calcChain.xml']:assert old.read(part)==z.read(part),part
 books=[E.fromstring(a.read('xl/workbook.xml')) for a in [old,z]]
 for book in books:
  for c in book.findall(n+'calcPr'):book.remove(c)
 assert E.tostring(books[0])==E.tostring(books[1])
 print(json.dumps({'passed':True,'sample':sample,'grandTotals':grand,'output':actual,'cacheUnchanged':True,'opaqueAndSourcePreserved':True}))
