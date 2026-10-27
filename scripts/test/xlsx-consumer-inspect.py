#!/usr/bin/env python3
"""Independent finite consumer audit: report differences, never infer no loss from a successful open.
Usage: xlsx-consumer-inspect.py SOURCE RETURNED. openpyxl is used only to translate explicit shared formula members.
Opaque embedded bytes and changed external dependencies stay unqualified unless a separate domain check proves them.
"""
import re,posixpath,urllib.parse
from decimal import Decimal, InvalidOperation
import sys,zipfile,xml.etree.ElementTree as E,hashlib,json,collections
from openpyxl.formula.translate import Translator
X='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
D='{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
def xml_semantics(node):return [node.tag.split('}')[-1],sorted(node.attrib.items()),node.text or '',[xml_semantics(n) for n in node]]
def shadow_key(node):
 attrs={'sx':'100000','sy':'100000','kx':'0','ky':'0',**node.attrib}
 return [sorted(attrs.items()),[xml_semantics(n) for n in node]]
def color(node):
 if node is None:return None
 fill=node.find(A+'solidFill')
 if node.find(A+'noFill') is not None:return ['none']
 if fill is None:return None
 return [[child.tag.split('}')[-1],sorted(child.attrib.items()),[[n.tag.split('}')[-1],sorted(n.attrib.items())] for n in child]] for child in fill]
def name_attributes(node):
 attrs=dict(node.attrib)
 for key in ['hidden','function','vbProcedure','xlm']:
  if attrs.get(key) in ['0','false']:attrs.pop(key)
  elif attrs.get(key)=='true':attrs[key]='1'
 return attrs
def header_sections(text):
 sections={'L':'','C':'','R':''};side='L';at=0
 while at<len(text):
  if text[at]!='&':sections[side]+=text[at];at+=1;continue
  if at+1>=len(text):return {'unqualified':text}
  code=text[at+1];at+=2
  if code=='&':sections[side]+='&'
  elif code in sections:side=code
  elif code=='"':
   end=text.find('"',at)
   if end<0:return {'unqualified':text}
   at=end+1
  elif code.isdigit():
   while at<len(text) and (text[at].isdigit() or text[at]=='.'):at+=1
  elif code=='K' and re.match('[0-9A-Fa-f]{6}',text[at:]):at+=6
  else:return {'unqualified':text}
 return sections
def drawing_shapes(z,names):
 shapes=[];anchors=[]
 for part in names:
  if not part.startswith('xl/drawings/') or '/_rels/' in part or not part.endswith('.xml'):continue
  root=E.fromstring(z.read(part))
  if root.tag!=D+'wsDr':continue
  for anchor in root:
   shape=anchor.find(D+'sp')
   if shape is None:continue
   props=shape.find(D+'nvSpPr/'+D+'cNvPr');geometry=shape.find(D+'spPr/'+A+'prstGeom');line=shape.find(D+'spPr/'+A+'ln');runs=[]
   for run in shape.findall(D+'txBody/'+A+'p/'+A+'r'):
    style=run.find(A+'rPr');font=style.find(A+'latin') if style is not None else None
    runs.append({'text':run.findtext(A+'t',''),'font':font.attrib.get('typeface') if font is not None else None,'size':style.attrib.get('sz') if style is not None else None,'bold':style.attrib.get('b','0') in ['1','true'] if style is not None else False,'italic':style.attrib.get('i','0') in ['1','true'] if style is not None else False,'color':color(style),'outline':color(style.find(A+'ln')) or ['none'] if style is not None else ['none'],'shadow':[shadow_key(n) for n in style.findall(A+'effectLst/'+A+'outerShdw')] if style is not None else []})
   shapes.append({'name':props.attrib.get('name') if props is not None else None,'altText':props.attrib.get('descr','') if props is not None else '', 'geometry':geometry.attrib.get('prst') if geometry is not None else 'custom','fill':color(shape.find(D+'spPr')),'outline':color(line),'runs':runs})
   origin=anchor.find(D+'from');extent=anchor.find(D+'ext');transform_extent=shape.find(D+'spPr/'+A+'xfrm/'+A+'ext')
   anchors.append({'name':props.attrib.get('name') if props is not None else None,'from':{n.tag.split('}')[-1]:n.text for n in origin} if origin is not None else None,'extent':dict(extent.attrib) if extent is not None else dict(transform_extent.attrib) if transform_extent is not None else None})
 return sorted(shapes,key=lambda n:n['name'] or ''),sorted(anchors,key=lambda n:n['name'] or '')
def decode(s):
 if s is None:return None
 return re.sub(r"_x([0-9A-Fa-f]{4})_",lambda m:chr(int(m[1],16)),s).encode("utf-16","surrogatepass").decode("utf-16","surrogatepass")
def visible_text(node):
 # Phonetic annotations are not visible rich-string runs.
 return ''.join(n.text or '' for n in node.findall(X+'t')+node.findall(X+'r/'+X+'t')) if node is not None else ''
def exact_number(raw):
 try:
  value=Decimal(raw)
  if not value.is_finite():return ['invalid',raw]
  sign,digits,exponent=value.as_tuple();digits=list(digits)
  if not any(digits):return [sign,[0],0]
  while digits[-1]==0:digits.pop();exponent+=1
  return [sign,digits,exponent]
 except (InvalidOperation,TypeError):return ['invalid',raw]
def normalized_formula(s,external):
 if not s:return s
 # Work only outside Excel string literals, and only on workbook indices followed by a sheet reference.
 chunks=re.split(r'("(?:[^"]|"")*")',s)
 for i in range(0,len(chunks),2):
  chunks[i]=re.sub(r"\[(\d+)\](?=(?:[^'!\[\]]+'|[A-Za-z_][\w.]*)!)",lambda m:'['+json.dumps(external.get(m[1],m[0]),ensure_ascii=False)+']',chunks[i])
  chunks[i]=re.sub(r"'([A-Za-z_][A-Za-z0-9_]*)'!",r'\1!',chunks[i])
 return ''.join(chunks)
def inspect(file):
 with zipfile.ZipFile(file) as z:
  names=z.namelist();book=E.fromstring(z.read('xl/workbook.xml'));rels={r.attrib['Id']:r.attrib['Target'] for r in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))};strings=[];external={}
  for i,r in enumerate(book.findall(X+"externalReferences/"+X+"externalReference"),1):
   target=rels[r.attrib[R+"id"]];part=target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/"+target);rp=posixpath.dirname(part)+"/_rels/"+posixpath.basename(part)+".rels"
   external[str(i)]=sorted(urllib.parse.unquote(n.attrib["Target"]) for n in E.fromstring(z.read(rp)) if n.attrib.get("TargetMode")=="External")
  def formula_key(s):return normalized_formula(s,external)
  if 'xl/sharedStrings.xml' in names:strings=[visible_text(si) for si in E.fromstring(z.read('xl/sharedStrings.xml'))]
  cells={};comment_owners={};print_settings={}
  for sheet in book.find(X+'sheets'):
   target=rels[sheet.attrib[R+'id']];part=target.lstrip('/') if target.startswith('/') else 'xl/'+target;root=E.fromstring(z.read(part));entries=root.findall(X+'sheetData/'+X+'row/'+X+'c');masters={}
   setup=root.find(X+'pageSetup');margins=root.find(X+'pageMargins');options=root.find(X+'printOptions');hf=root.find(X+'headerFooter')
   print_settings[sheet.attrib['name']]={'explicit':any(n is not None for n in [setup,margins,options,hf]),'setup':{k:setup.attrib.get(k,v) if setup is not None else v for k,v in [('paperSize','1'),('orientation','portrait'),('scale','100')]},'margins':{k:exact_number(v) for k,v in margins.attrib.items()} if margins is not None else None,'gridlines':options.attrib.get('gridLines','0') in ['1','true'] if options is not None else False,'headers':{n.tag.split('}')[-1]:n.text or '' for n in hf} if hf is not None else {}}
   rp=posixpath.dirname(part)+'/_rels/'+posixpath.basename(part)+'.rels'
   if rp in names:
    for edge in E.fromstring(z.read(rp)):
     if edge.attrib.get('Type','').endswith('/comments') and edge.attrib.get('TargetMode')!='External':
      target=edge.attrib['Target'];comment=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.dirname(part)+'/'+target);comment_owners.setdefault(comment,[]).append(sheet.attrib['name'])
   for c in entries:
    f=c.find(X+'f')
    if f is not None and f.attrib.get('t')=='shared' and f.text:masters[f.attrib['si']]=(c.attrib['r'],f.text)
   for c in entries:
    f=c.find(X+'f');v=c.find(X+'v');raw=v.text if v is not None else None;kind=c.attrib.get('t','n');value=raw
    if kind=='s':value=decode(strings[int(raw)])
    elif kind=='inlineStr':value=decode(visible_text(c.find(X+'is')))
    elif kind=='str':value=decode(raw)
    elif kind=='b':value=raw=='1'
    elif kind=='n' and raw is not None:
     try:value=float(raw)
     except ValueError:pass
    formula=f.text if f is not None else None
    if f is not None and f.attrib.get('t')=='shared':
     origin,text=masters[f.attrib['si']];formula=Translator('='+text,origin=origin).translate_formula(c.attrib['r'])[1:]
    cells[sheet.attrib['name']+'!'+c.attrib['r']]={'value':value,'type':'text' if kind in ['s','inlineStr','str'] else kind,'formula':formula_key(formula),'style':c.attrib.get('s','0'),'exactValue':exact_number(raw) if kind=='n' and raw is not None else value}
  C='{http://schemas.openxmlformats.org/drawingml/2006/chart}'
  charts=[];chart_numbers=[]
  for part in names:
   if not part.startswith('xl/charts/') or not part.endswith('.xml') or '/_rels/' in part:continue
   chart=E.fromstring(z.read(part))
   if chart.tag!=C+'chartSpace':continue
   kind='line' if chart.find('.//'+C+'lineChart') is not None else None
   bar=chart.find('.//'+C+'barChart')
   if bar is not None:kind='column' if bar.find(C+'barDir').attrib.get('val')=='col' else 'bar'
   refs=[re.sub(r"^'([A-Za-z_][A-Za-z0-9_]*)'!",r"\1!",n.text or '') for n in chart.iter(C+'f')]
   charts.append({'kind':kind,'references':refs,'categories':[n.text for n in chart.findall('.//'+C+'cat/'+C+'strRef/'+C+'strCache/'+C+'pt/'+C+'v')],'values':[float(n.text) for n in chart.findall('.//'+C+'val/'+C+'numRef/'+C+'numCache/'+C+'pt/'+C+'v')]})
   chart_numbers.append([exact_number(n.text) for n in chart.findall('.//'+C+'numCache/'+C+'pt/'+C+'v')])
  cache_sources={}
  for cache in book.findall(X+'pivotCaches/'+X+'pivotCache'):
   target=rels[cache.attrib[R+'id']];part=target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
   root=E.fromstring(z.read(part));source=root.find(X+'cacheSource');ws=source.find(X+'worksheetSource') if source is not None else None
   cache_sources[cache.attrib['cacheId']]={'type':source.attrib.get('type') if source is not None else None,'worksheetSource':dict(ws.attrib) if ws is not None else None,'fields':[n.attrib.get('name') for n in root.findall(X+'cacheFields/'+X+'cacheField')]}
  pivots=[];pivot_sources=[]
  for part in names:
   if not part.startswith('xl/pivotTables/') or not part.endswith('.xml') or '/_rels/' in part:continue
   pivot=E.fromstring(z.read(part));data=pivot.find(X+'dataFields');pivots.append({'name':pivot.attrib['name'],'range':pivot.find(X+'location').attrib['ref'],'aggregates':[(f.attrib.get('fld'),f.attrib.get('subtotal','sum')) for f in data] if data is not None else []})
   pivot_sources.append(cache_sources.get(pivot.attrib.get('cacheId')))
  shapes,anchors=drawing_shapes(z,names)
  themes=[]
  for edge in E.fromstring(z.read('xl/_rels/workbook.xml.rels')):
   if not edge.attrib.get('Type','').endswith('/theme') or edge.attrib.get('TargetMode')=='External':continue
   target=edge.attrib['Target'];part=target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target);theme=E.fromstring(z.read(part));elements=theme.find(A+'themeElements')
   themes.append({'colors':[[n.tag.split('}')[-1],[[c.tag.split('}')[-1],sorted(c.attrib.items())] for c in n]] for n in elements.find(A+'clrScheme')],'fonts':{name:elements.find(A+'fontScheme/'+A+name+'/'+A+'latin').attrib.get('typeface') for name in ['majorFont','minorFont']}})
  notes=[]
  for part in names:
   if not re.match(r'^xl/(?:comments[^/]*|comments/[^/]+)\.xml$',part):continue
   root=E.fromstring(z.read(part))
   if root.tag!=X+'comments':continue
   authors=[n.text or '' for n in root.findall(X+'authors/'+X+'author')]
   notes.extend({'sheets':sorted(comment_owners.get(part,[])),'ref':n.attrib['ref'],'author':authors[int(n.attrib['authorId'])],'text':decode(visible_text(n.find(X+'text')))} for n in root.findall(X+'commentList/'+X+'comment'))
  return {'cells':cells,'themes':themes,'printSettings':print_settings,'shapes':shapes,'shapeAnchors':anchors,'notes':sorted(notes,key=lambda n:(n['ref'],n['author'],n['text'])),'charts':charts,'chartNumbers':chart_numbers,'pivots':pivots,'pivotSources':pivot_sources,'externalTargets':sorted(external.values()),'sheets':[(s.attrib['name'],s.attrib.get('state','visible')) for s in book.find(X+'sheets')],'names':sorted(json.dumps([name_attributes(n),formula_key(n.text)],sort_keys=True,ensure_ascii=False) for n in book.findall(X+'definedNames/'+X+'definedName')),'images':collections.Counter(hashlib.sha256(z.read(n)).hexdigest() for n in names if n.startswith('xl/media/') and not n.endswith('/')),'embeddings':collections.Counter(hashlib.sha256(z.read(n)).hexdigest() for n in names if n.startswith('xl/embeddings/') and not n.endswith('/')),'externalLinkParts':sum(n.startswith('xl/externalLinks/') for n in names),'parts':len(names),'date1904':str(int(book.find(X+'workbookPr').attrib.get('date1904','0') in ['1','true'])) if book.find(X+'workbookPr') is not None else '0'}
def compare(a,b):
 changes=[];missing=[];formula_changes=[];cache_changes=0;binary_equivalent_decimal_changes=[]
 for ref,c in a['cells'].items():
  d=b['cells'].get(ref)
  if d is None:
   if c['value'] is not None or c['formula'] is not None:missing.append(ref)
   continue
  if c['formula']!=d['formula']:formula_changes.append(ref)
  different=c['exactValue']!=d['exactValue'] or c['type']!=d['type']
  if c['formula'] is None and different:changes.append(ref)
  if c['formula'] is None and c['type']==d['type']=='n' and c['exactValue']!=d['exactValue'] and c['value']==d['value']:binary_equivalent_decimal_changes.append(ref)
  if c['formula'] is not None and different:cache_changes+=1
 added=[ref for ref,c in b['cells'].items() if ref not in a['cells'] and (c['value'] is not None or c['formula'] is not None)]
 report={'scope':'Finite cell/formula, sheet/name, chart data/reference, pivot source/aggregate, external target, image/embedding bytes and date-system comparison. Formula caches may recalculate. Styles, layout, opaque dependency semantics and all consumer behavior remain unqualified.','sourceCells':len(a['cells']),'chartsOriginal':a['charts'],'chartsReturned':b['charts'],'chartDataAndReferencesRetained':a['charts']==b['charts'] and a['chartNumbers']==b['chartNumbers'],'pivotsOriginal':a['pivots'],'pivotsReturned':b['pivots'],'pivotDefinitionsRetained':a['pivots']==b['pivots'],'pivotSourceReferencesRetained':a['pivotSources']==b['pivotSources'],'externalTargetsRetained':a['externalTargets']==b['externalTargets'],'returnedCells':len(b['cells']),'unexpectedNonblankCells':len(added),'nonblankCellsMissing':len(missing),'literalValueOrTypeChanges':len(changes),'authoredFormulaChanges':len(formula_changes),'formulaCacheChanges':cache_changes,'sheetIdentityRetained':a['sheets']==b['sheets'],'namesRetained':a['names']==b['names'],'imagesOriginal':sum(a['images'].values()),'imagesReturned':sum(b['images'].values()),'imagesByteRetained':a['images']==b['images'],'embeddedOriginal':sum(a['embeddings'].values()),'embeddedReturned':sum(b['embeddings'].values()),'embeddedByteRetained':a['embeddings']==b['embeddings'],'externalPartsOriginal':a['externalLinkParts'],'externalPartsReturned':b['externalLinkParts'],'dateSystemRetained':a['date1904']==b['date1904'],'changedCoordinates':{'missing':missing[:20],'added':added[:20],'values':changes[:20],'formulas':formula_changes[:20]}}
 report['passed']=not missing and not added and not changes and not formula_changes and all(report[key] for key in ['sheetIdentityRetained','namesRetained','imagesByteRetained','embeddedByteRetained','dateSystemRetained','chartDataAndReferencesRetained','pivotDefinitionsRetained','pivotSourceReferencesRetained','externalTargetsRetained'])
 report['binary64EquivalentDecimalChanges']=len(binary_equivalent_decimal_changes)
 report['numericComparison']='Exact decimal/sign changes are reported separately even when binary64 values compare equal; these alone do not establish user-visible numeric loss.'
 report.update({'shapesOriginal':a['shapes'],'shapesReturned':b['shapes'],'shapeTextGeometryStyleRetained':a['shapes']==b['shapes'],'shapeAnchorsOriginal':a['shapeAnchors'],'shapeAnchorsReturned':b['shapeAnchors'],'shapeAnchorGeometryRetained':a['shapeAnchors']==b['shapeAnchors'],'legacyNotesRetained':a['notes']==b['notes']})
 report['scope']+=' Simple DrawingML text/geometry/run styles and legacy-note text/authors are checked; groups, alternate-content drawings, full typography and physical placement/layout remain unqualified.'
 report['passed']=report['passed'] and report['shapeTextGeometryStyleRetained'] and report['legacyNotesRetained']
 def print_key(value):return {k:({n:header_sections(text) for n,text in v.items()} if k=='headers' else v) for k,v in value.items() if k!='explicit'}
 report.update({'themeColorsAndFontsRetained':not a['themes'] or a['themes']==b['themes'],'explicitThemesOriginal':len(a['themes']),'explicitThemesReturned':len(b['themes']),'implicitThemeAddedByConsumer':not a['themes'] and bool(b['themes']),'printSettingsOriginal':a['printSettings'],'printSettingsReturned':b['printSettings'],'basicPrintSettingsRetained':all(sheet in b['printSettings'] and print_key(value)==print_key(b['printSettings'][sheet]) for sheet,value in a['printSettings'].items() if value['explicit'])})
 report['scope']+=' Theme/font catalogs and print settings cover explicitly authored source fields; consumer default-theme creation is reported separately. Header checks cover section text, not typography.'
 report['passed']=report['passed'] and report['themeColorsAndFontsRetained'] and report['basicPrintSettingsRetained']
 return report
if __name__=='__main__':
 print(json.dumps(compare(inspect(sys.argv[1]),inspect(sys.argv[2])),ensure_ascii=False))
