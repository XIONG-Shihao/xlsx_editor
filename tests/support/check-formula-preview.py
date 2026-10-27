"""Read-only SQLite/source oracle. Never export or evaluate a workbook projection."""
import hashlib,json,math,sqlite3,struct,sys,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

def file_record(path,root=None):
    data=path.read_bytes()
    return {'path':str(path.relative_to(root)) if root else str(path),'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()}

def scalar(value):
    if isinstance(value,bytes): return {'type':'blob','hex':value.hex()}
    if isinstance(value,float): return {'type':'float64','bits':struct.pack('>d',value).hex()}
    if value is None: return {'type':'null'}
    return {'type':'integer' if isinstance(value,int) else 'text','value':value}

def snapshot(directory):
    root=Path(directory).resolve(strict=True)
    if not root.is_dir(): raise AssertionError('Expected owned index directory')
    files=[];databases=[];excluded=[]
    for path in sorted(root.rglob('*')):
        if path.is_symlink(): raise AssertionError('Oracle never follows index symlinks')
        if not path.is_file(): continue
        if path.name.endswith('.sqlite-shm'):
            excluded.append(str(path.relative_to(root)))
            continue
        files.append(file_record(path,root))
        if path.name!='cells.sqlite': continue
        before=file_record(path)
        db=sqlite3.connect(path.as_uri()+'?mode=ro',uri=True)
        try:
            db.execute('BEGIN')
            schema=db.execute("SELECT name,type,sql FROM sqlite_schema WHERE type IN ('table','index','trigger','view') ORDER BY name,type").fetchall()
            tables=[];total=0
            for name,kind,sql in schema:
                if kind!='table': continue
                quoted='"'+name.replace('"','""')+'"'
                columns=[r[1] for r in db.execute('PRAGMA table_info('+quoted+')')]
                rows=[]
                for row in db.execute('SELECT * FROM '+quoted):
                    encoded=json.dumps([scalar(v) for v in row],ensure_ascii=False,sort_keys=True,separators=(',',':')).encode()
                    total+=len(encoded)
                    if total>64*1024*1024 or len(rows)>=200000: raise AssertionError('Independent SQL oracle budget exceeded')
                    rows.append(encoded)
                digest=hashlib.sha256()
                for row in sorted(rows): digest.update(len(row).to_bytes(8,'big'));digest.update(row)
                table={'name':name,'columns':columns,'rows':len(rows),'sha256':digest.hexdigest()}
                if name in ('versions','version_ancestors','metadata','calculation_policy'):
                    table['records']=[json.loads(row) for row in sorted(rows)]
                tables.append(table)
            db.rollback()
        finally: db.close()
        assert file_record(path)==before,'Read-only oracle altered SQLite bytes'
        databases.append({'path':str(path.relative_to(root)),'schema':schema,'tables':tables,'logicalBytes':total})
    return {'schema':'workbook.xlsx.formula-preview-persistent-sql.v1','files':files,'databases':databases,
            'excluded':{'sqliteShm':excluded,'reason':'Advisory shared-memory lock/read-mark bytes only; main DB and WAL bytes remain included. SQLite connection-private TEMP files outside this accepted index are not treated as workbook persistence.'}}

def source(path):
    n={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with zipfile.ZipFile(path) as archive:
        assert archive.testzip() is None
        members=[{'path':i.filename,'bytes':i.file_size,'sha256':hashlib.sha256(archive.read(i.filename)).hexdigest()} for i in sorted(archive.infolist(),key=lambda x:x.filename)]
        workbook=ET.fromstring(archive.read('xl/workbook.xml'))
        cells={}
        for item in members:
            if item['path'].startswith('xl/worksheets/sheet') and item['path'].endswith('.xml'):
                root=ET.fromstring(archive.read(item['path'])); records={}
                for cell in root.findall('.//s:sheetData/s:row/s:c',n):
                    ref=cell.attrib['r'];assert ref not in records
                    formula=cell.find('s:f',n);value=cell.find('s:v',n)
                    records[ref]={'attributes':cell.attrib,'formula':None if formula is None else formula.text,'formulaAttributes':None if formula is None else formula.attrib,'rawCache':None if value is None else value.text,'inline':''.join(cell.itertext()) if cell.attrib.get('t')=='inlineStr' else None}
                cells[item['path']]=records
        names=[{'attributes':x.attrib,'formula':x.text} for x in workbook.findall('s:definedNames/s:definedName',n)]
        policy=workbook.find('s:calcPr',n)
    return {'schema':'workbook.xlsx.formula-preview-source.v1','file':file_record(Path(path)),'members':members,'cells':cells,'names':names,'calcPr':None if policy is None else policy.attrib}

if __name__=='__main__':
    assert len(sys.argv)==3 and sys.argv[1] in ('--snapshot','--source')
    result=snapshot(sys.argv[2]) if sys.argv[1]=='--snapshot' else source(sys.argv[2])
    print(json.dumps(result,ensure_ascii=False,allow_nan=False,separators=(',',':')))
