#!/usr/bin/env python3
"""Independent OPC/cache/display oracle; never imports the production pivot model."""
import argparse,collections,hashlib,json,pathlib,posixpath,zipfile,xml.etree.ElementTree as E
from openpyxl import load_workbook
p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('saved');p.add_argument('--aggregate',choices=['sum','count'],default='sum');p.add_argument('--source-amount',type=int,default=2);p.add_argument('--allow-stale',action='store_true');p.add_argument('--output');a=p.parse_args()
S='http://schemas.openxmlformats.org/spreadsheetml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';P='http://schemas.openxmlformats.org/package/2006/relationships';T='http://schemas.openxmlformats.org/package/2006/content-types'
tag=lambda name:'{'+S+'}'+name;sha=lambda b:hashlib.sha256(b).hexdigest()
def parts(file):
    with zipfile.ZipFile(file)as z:
        assert len(z.namelist())==len(set(z.namelist()));return {name:z.read(name)for name in z.namelist()}
before=parts(a.source);after=parts(a.saved)
def xml(file,ns,root):
    value=E.fromstring(after[file]);assert value.tag=='{'+ns+'}'+root;return value
def relfile(file):return posixpath.join(posixpath.dirname(file),'_rels',posixpath.basename(file)+'.rels')
def target(file,value):
    result=posixpath.normpath(posixpath.join(posixpath.dirname(file),value));assert not result.startswith('../')and result in after;return result
def relation(file,kind):
    root=xml(relfile(file),P,'Relationships');values=[n for n in root if n.get('Type')==R+'/'+kind];assert len(values)==1 and not values[0].get('TargetMode');return values[0],target(file,values[0].get('Target'))
workbook=xml('xl/workbook.xml',S,'workbook');wb_rels=xml('xl/_rels/workbook.xml.rels',P,'Relationships');ids={n.get('Id'):n for n in wb_rels};assert len(ids)==len(wb_rels)
sheets=workbook.find(tag('sheets'));names=[n.get('name')for n in sheets];assert names==['Sheet1','Summary']
paths={n.get('name'):target('xl/workbook.xml',ids[n.get('{'+R+'}id')].get('Target'))for n in sheets}
summary=xml(paths['Summary'],S,'worksheet');assert summary.find(tag('pivotTableParts')) is None,'Pivot linkage belongs to OPC relationships, not worksheet markup'
edge,table_file=relation(paths['Summary'],'pivotTable');assert edge.get('Id')=='rIdPivot';table=xml(table_file,S,'pivotTableDefinition');assert table.get('name')=='SalesPivot'and table.get('cacheId')=='1'
edge,cache_file=relation(table_file,'pivotCacheDefinition');cache=xml(cache_file,S,'pivotCacheDefinition');caches=workbook.find(tag('pivotCaches'));assert len(caches)==1 and caches[0].get('cacheId')=='1';cache_edge=ids[caches[0].get('{'+R+'}id')];assert cache_edge.get('Type')==R+'/pivotCacheDefinition'and target('xl/workbook.xml',cache_edge.get('Target'))==cache_file
edge,record_file=relation(cache_file,'pivotCacheRecords');assert edge.get('Id')==cache.get('{'+R+'}id');records=xml(record_file,S,'pivotCacheRecords')
source=cache.find(tag('cacheSource'));assert source.get('type')=='worksheet';ws=source.find(tag('worksheetSource'));assert ws.attrib=={'ref':'A1:C5','sheet':'Sheet1'};assert cache.get('refreshOnLoad')=='0'and cache.get('recordCount')=='4'and records.get('count')=='4'and len(records)==4
fields=cache.find(tag('cacheFields'));assert fields.get('count')=='3'and len(fields)==3;assert [f.get('name')for f in fields]==['Category','Amount','Flag'];shared=[]
for field in fields:
    items=field.find(tag('sharedItems'));assert int(items.get('count'))==len(items);values=[]
    for item in items:
        kind=item.tag.removeprefix('{'+S+'}')
        if kind=='m':assert not item.attrib;value=None
        elif kind=='s':value=item.get('v')
        elif kind=='b':assert item.get('v')in('0','1');value=item.get('v')=='1'
        elif kind=='n':value=int(item.get('v'))
        else:raise AssertionError('Unqualified shared item '+kind)
        values.append(value)
    office_types={('number' if type(v) is int else 'string') for v in values if v is not None}
    assert items.get('containsString')==str(int('string' in office_types))
    assert items.get('containsSemiMixedTypes')==str(int('string' in office_types or None in values))
    assert items.get('containsMixedTypes')==str(int(len(office_types)>1))
    assert len({(type(v).__name__,v)for v in values})==len(values);shared.append(values)
decoded=[]
for row in records:
    assert row.tag==tag('r')and len(row)==3;values=[]
    for c,item in enumerate(row):
        assert item.tag==tag('x')and item.get('v')==str(int(item.get('v')));index=int(item.get('v'));assert 0<=index<len(shared[c]);values.append(shared[c][index])
    decoded.append(values)
assert decoded[1]==['B',5,False]and decoded[2]==['A',-1,True]and decoded[3]==['A',None,False];assert decoded[0][0]=='A'and decoded[0][2]is True and type(decoded[0][1])is int
data_fields=table.find(tag('dataFields'));assert data_fields.get('count')=='1'and len(data_fields)==1;field=data_fields[0];assert field.get('subtotal')==a.aggregate and field.get('fld')=='1';row_fields=table.find(tag('rowFields'));assert row_fields.get('count')=='1'and row_fields[0].get('x')=='0'
groups=collections.OrderedDict()
for label,value,flag in decoded:groups.setdefault(label,0);groups[label]+=0 if value is None else(1 if a.aggregate=='count'else value)
caption=('Sum of 'if a.aggregate=='sum'else'Count of ')+'Amount';expected=[('Category',caption),*groups.items(),('Grand Total',sum(groups.values()))]
book=load_workbook(a.saved,data_only=False);assert len(book['Summary']._pivots)==1 and book['Summary']._pivots[0].name=='SalesPivot';assert list(book['Summary'].values)==expected
assert len(table.find(tag('pivotFields'))[0].find(tag('items')))==2
columns=table.find(tag('colItems'));assert columns.get('count')=='1' and len(columns)==1 and len(columns[0])==0
assert table.find(tag('location')).get('firstHeaderRow')=='1'
assert table.find(tag('location')).get('ref')=='A1:B4';assert len(table.find(tag('rowItems')))==3;assert len(table.find(tag('pivotFields')))==3
source_rows=list(book['Sheet1'].values);assert source_rows==[('Category','Amount','Flag'),('A',a.source_amount,True),('B',5,False),('A',-1,True),('A',None,False)]
matches_cache=decoded[0][1]==a.source_amount;assert matches_cache or a.allow_stale,'Source differs from saved cache without explicit manual-refresh scope'
types=xml('[Content_Types].xml',T,'Types');mimes={n.get('PartName','').removeprefix('/'):n.get('ContentType')for n in types if n.tag=='{'+T+'}Override'}
for file,kind in [(table_file,'pivotTable'),(cache_file,'pivotCacheDefinition'),(record_file,'pivotCacheRecords')]:assert mimes[file]=='application/vnd.openxmlformats-officedocument.spreadsheetml.'+kind+'+xml'
changed={name for name in before if name in after and before[name]!=after[name]};added=set(after)-set(before);removed=set(before)-set(after);assert not removed
allowed={paths['Sheet1'],paths['Summary'],table_file,cache_file,record_file,'xl/workbook.xml','xl/_rels/workbook.xml.rels','[Content_Types].xml'};assert changed<=allowed,(changed-allowed)
allowed_added={paths['Summary'],table_file,cache_file,record_file,relfile(paths['Summary']),relfile(table_file),relfile(cache_file)};assert added<=allowed_added,(added-allowed_added)
preserved={name:sha(value)for name,value in before.items()if name not in changed};assert all(after[name]==before[name]for name in preserved);assert after.get('customXml/opaque.bin')==before.get('customXml/opaque.bin')
report={'status':'pass','sourceSha256':sha(pathlib.Path(a.source).read_bytes()),'savedSha256':sha(pathlib.Path(a.saved).read_bytes()),'pivotName':'SalesPivot','cacheRecords':decoded,'cachedRows':[list(row)for row in expected],'aggregate':a.aggregate,'sourceMatchesCache':matches_cache,'manualRefreshOnly':True,'recognizedBy':'openpyxl '+__import__('openpyxl').__version__,'graph':{'worksheet':paths['Summary'],'table':table_file,'cache':cache_file,'records':record_file},'changedParts':sorted(changed),'addedParts':sorted(added),'preservedParts':preserved,'scope':'Independent raw OPC/index/type/count/cache/display and openpyxl PivotTable readback. No Excel runtime implied.'}
if a.output:pathlib.Path(a.output).write_text(json.dumps(report,indent=2)+'\n')
print(json.dumps(report))
