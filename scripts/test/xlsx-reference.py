#!/usr/bin/env python3
"""Independent OOXML inventory and hand-authored literal answer-key checks (stdlib)."""
import argparse, datetime, hashlib, json, math, platform, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
X='http://schemas.openxmlformats.org/spreadsheetml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';P='http://schemas.openxmlformats.org/package/2006/relationships'
NS={'x':X,'r':R,'p':P}
def sha(file):return hashlib.sha256(Path(file).read_bytes()).hexdigest()
def revision_identity(root):
    tracked=['fixtures/build-corpus.py','fixtures/expectations.json','fixtures/budgets.json','scripts/test/xlsx-reference.py']
    files={name:sha(root/name) for name in tracked if (root/name).is_file()}
    return {'files':files,'sha256':hashlib.sha256(json.dumps(files,sort_keys=True).encode()).hexdigest(),'build':{'kind':'source-only','bundleHash':None,'reason':'No packaged app or GUI executed by this harness.'}}
def inspect(file,include_cells=False):
    with zipfile.ZipFile(file) as archive: parts={name:archive.read(name) for name in archive.namelist()}
    xml=lambda name:ET.fromstring(parts[name]);book=xml('xl/workbook.xml')
    rels={r.attrib['Id']:r.attrib['Target'] for r in xml('xl/_rels/workbook.xml.rels') if r.attrib.get('TargetMode')!='External'}
    strings=[''.join(n.text or '' for n in si.iter(f'{{{X}}}t')) for si in xml('xl/sharedStrings.xml')] if 'xl/sharedStrings.xml' in parts else []
    sheets=[];cells={};formulas=0;nonempty=0
    for sheet in book.findall('x:sheets/x:sheet',NS):
      target=rels[sheet.attrib[f'{{{R}}}id']];part=target.lstrip('/') if target.startswith('/') else 'xl/'+target
      root=xml(part);entries=root.findall('.//x:sheetData/x:row/x:c',NS);used=0;formula_count=0
      for c in entries:
        v=c.find('x:v',NS);f=c.find('x:f',NS);raw=v.text if v is not None else None;t=c.attrib.get('t','n')
        value=raw
        if t=='inlineStr':value=''.join(n.text or '' for n in c.findall('.//x:t',NS))
        elif t=='s':value=strings[int(raw)]
        elif t=='b':value=raw=='1'
        elif t=='n' and raw is not None:
          try:value=int(raw) if raw.lstrip('-').isdigit() else float(raw)
          except ValueError:pass
        if raw is not None or f is not None or t=='inlineStr':used+=1
        if f is not None:formula_count+=1
        if include_cells:cells[sheet.attrib['name']+'!'+c.attrib['r']]={'value':value,'rawValue':raw,'formula':f.text if f is not None else None,'type':t,'error':raw if t=='e' else None,'style':c.attrib.get('s')}
      nonempty+=used;formulas+=formula_count
      sheets.append({'name':sheet.attrib['name'],'part':part,'state':sheet.attrib.get('state','visible'),'usedCells':used,'storedCells':len(entries),'formulas':formula_count,'merges':len(root.findall('.//x:mergeCell',NS)),'validations':len(root.findall('.//x:dataValidation',NS)),'protected':root.find('x:sheetProtection',NS) is not None})
    properties=book.find('x:workbookPr',NS)
    result={'file':str(file),'sha256':sha(file),'bytes':Path(file).stat().st_size,'dateSystem':'1904' if properties is not None and properties.attrib.get('date1904') in ['1','true'] else '1900','usedCells':nonempty,'formulas':formulas,'sheets':sheets,'namedRanges':len(book.findall('x:definedNames/x:definedName',NS)),'workbookProtected':book.find('x:workbookProtection',NS) is not None,'features':{kind:sum(name.startswith(prefix) and not '/_rels/' in name and not name.endswith('/') for name in parts) for kind,prefix in [('tables','xl/tables/'),('charts','xl/charts/'),('imagesOrMedia','xl/media/'),('pivots','xl/pivotTables/'),('externalLinks','xl/externalLinks/'),('embeddings','xl/embeddings/'),('customXml','customXml/') ]},'members':{name:{'sha256':hashlib.sha256(value).hexdigest(),'bytes':len(value)} for name,value in sorted(parts.items())}}
    if include_cells:result['cells']=cells
    return result

def check_literals(inventory,key,calculated=False):
    checks=[]
    for target,expected in key.get('cells',{}).items():
      actual=inventory.get('cells',{}).get(target);failures=[]
      if actual is None:
        if not (calculated and expected.get('mayOmitBlankCell') and expected.get('value') is None):failures.append('missing cell')
      else:
        for prop in ['value','rawValue']:
          if prop in expected and not (calculated and prop=='rawValue') and actual[prop]!=expected[prop]:failures.append(prop)
        if not calculated and 'formula' in expected and actual['formula']!=expected['formula']:failures.append('formula')
        if not calculated and 'storedCache' in expected and actual['value']!=expected['storedCache']:failures.append('storedCache')
        if calculated and 'calculated' in expected:
          value=actual['value'];wanted=expected['calculated']
          if not isinstance(value,(int,float)) or not math.isclose(value,wanted,rel_tol=1e-12,abs_tol=1e-9):failures.append('calculated')
        if calculated and 'error' in expected and actual['error']!=expected['error']:failures.append('error')
        if calculated and expected.get('status') in ['unresolved','circular'] and not actual['error']:failures.append('explicit error required')
      checks.append({'target':target,'expected':expected,'actual':actual,'passed':not failures,'mismatches':failures,**({'referenceExclusion':'Exact stored numeric text uses the direct OOXML literal key; reference-engine normalization is recorded but is not an oracle for byte/value fidelity.'} if calculated and 'rawValue' in expected else {})})
    return checks

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--corpus',type=Path,required=True);parser.add_argument('--output',type=Path,required=True);parser.add_argument('--extra',action='append',default=[]);args=parser.parse_args()
    root=Path(__file__).resolve().parents[2];expected=json.loads((root/'fixtures/expectations.json').read_text());ident=revision_identity(root);files=[]
    for name,key in expected['fixtures'].items():
      inventory=inspect(args.corpus/(name+'.xlsx'),True);checks=check_literals(inventory,key);inventory.pop('cells');inventory['literalChecks']=checks;inventory['id']=name;inventory['provenance']='deterministic synthetic OOXML, not production representative';files.append(inventory)
    for extra in args.extra:
      file=Path(extra)
      if file.is_file():
        item=inspect(file);item['provenance']='existing local file; no independent cell answer key and not edited by this harness';files.append(item)
    result={'schema':'workbook.xlsx.corpus-evidence.v1','measuredAt':datetime.datetime.now(datetime.timezone.utc).isoformat(),'source':ident,'machine':{'platform':platform.platform(),'machine':platform.machine(),'python':platform.python_version()},'fixtures':files}
    result['sourceStableDuringRun']=ident==revision_identity(root)
    args.output.parent.mkdir(parents=True,exist_ok=True);args.output.write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n');print(json.dumps({'output':str(args.output),'fixtures':len(files),'literalFailures':sum(not c['passed'] for f in files for c in f.get('literalChecks',[]))}))
if __name__=='__main__':main()
