#!/usr/bin/env python3
"""Deterministic, synthetic OOXML fixtures; never imports a workbook implementation."""
import argparse, base64, hashlib, json, random, zipfile
from pathlib import Path
from xml.sax.saxutils import escape
X='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'

def cell(ref, value=None, formula=None, style=None, raw=None, error=None):
    attrs=f' r="{ref}"'+(f' s="{style}"' if style is not None else '')
    if formula is not None:
        return f'<c{attrs}'+(' t="e"' if error else '')+'><f>'+escape(formula)+'</f>'+('' if value is None and error is None else '<v>'+escape(str(error if error else value))+'</v>')+'</c>'
    if isinstance(value,str): return f'<c{attrs} t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'
    if isinstance(value,bool): return f'<c{attrs} t="b"><v>{int(value)}</v></c>'
    return f'<c{attrs}>'+('' if value is None and raw is None else '<v>'+str(raw if raw is not None else value)+'</v>')+'</c>'
def rows(*values): return ''.join(f'<row r="{i}">{"".join(v)}</row>' for i,v in enumerate(values,1))
def worksheet(body,extra=''): return f'<worksheet xmlns="{X}" xmlns:r="{R}"><sheetData>{body}</sheetData>{extra}</worksheet>'
def rels(entries): return f'<Relationships xmlns="{P}">'+''.join(f'<Relationship Id="{i}" Type="{R}/{t}" Target="{target}"/>' for i,t,target in entries)+'</Relationships>'
STYLES=f'<styleSheet xmlns="{X}"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.00"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
def workbook(sheets,date1904=False,names=None,protected=False,extras=None):
    parts={'_rels/.rels':rels([('office','officeDocument','xl/workbook.xml')]),'xl/styles.xml':STYLES}
    book=f'<workbook xmlns="{X}" xmlns:r="{R}"><workbookPr date1904="{int(date1904)}"/>'
    if protected: book+='<workbookProtection lockStructure="1"/>'
    book+='<sheets>'+''.join(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="sheet{i}" state="{state}"/>' for i,(name,_,state) in enumerate(sheets,1))+'</sheets>'
    if names: book+='<definedNames>'+''.join(f'<definedName name="{name}">{escape(formula)}</definedName>' for name,formula in names.items())+'</definedNames>'
    parts['xl/workbook.xml']=book+'<calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>'
    parts['xl/_rels/workbook.xml.rels']=rels([(f'sheet{i}','worksheet',f'worksheets/sheet{i}.xml') for i in range(1,len(sheets)+1)]+[('styles','styles','styles.xml')])
    overrides=[('xl/workbook.xml','spreadsheetml.sheet.main'),('xl/styles.xml','spreadsheetml.styles')]
    for i,(_,xml,_) in enumerate(sheets,1):
        parts[f'xl/worksheets/sheet{i}.xml']=xml
        overrides.append((f'xl/worksheets/sheet{i}.xml','spreadsheetml.worksheet'))
    parts.update(extras or {})
    for name in parts:
        if name.startswith('xl/tables/'): overrides.append((name,'spreadsheetml.table'))
        elif name.startswith('xl/charts/'): overrides.append((name,'drawingml.chart'))
        elif name.startswith('xl/drawings/') and not '/_rels/' in name: overrides.append((name,'drawing'))
    parts['[Content_Types].xml']=f'<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="bin" ContentType="application/octet-stream"/>'+''.join(f'<Override PartName="/{name}" ContentType="application/vnd.openxmlformats-officedocument.{kind}+xml"/>' for name,kind in overrides)+'</Types>'
    return parts

def corpus():
    rates=worksheet(rows([cell('A1','Tax rate'),cell('B1',0.09)]))
    pricing=worksheet(rows(
      [cell(f'{c}1',v) for c,v in zip('ABCDEF',['SKU','Quantity','Unit Price','Subtotal','Tax','Grand Total'])],
      [cell('A2','00123'),cell('B2',2),cell('C2',125,style=2),cell('D2',250,'B2*C2'),cell('E2',0.09,'TaxRate'),cell('F2',272.5,'D2*(1+E2)',2)],
      [cell('A3','00456'),cell('B3',1),cell('C3',80),cell('D3',80,'B3*C3'),cell('E3',0.09,'TaxRate'),cell('F3',87.2,'D3*(1+E3)',2)],
      [cell('A4','Total'),cell('D4',330,'SUM(D2:D3)'),cell('F4',359.7,'SUM(F2:F3)',2)],
      [cell('A5','Table subtotal'),cell('D5',330,'SUM(PricingItems[Subtotal])')]), '<tableParts count="1"><tablePart r:id="table1"/></tableParts>')
    table=f'<table xmlns="{X}" id="1" name="PricingItems" displayName="PricingItems" ref="A1:F3" totalsRowShown="0"><autoFilter ref="A1:F3"/><tableColumns count="6">'+''.join(f'<tableColumn id="{i}" name="{name}"/>' for i,name in enumerate(['SKU','Quantity','Unit Price','Subtotal','Tax','Grand Total'],1))+'</tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>'
    yield 'pricing-schedule',workbook([('Pricing',pricing,'visible'),('Rates',rates,'hidden')],names={'TaxRate':'Rates!$B$1'},extras={'xl/tables/table1.xml':table,'xl/worksheets/_rels/sheet1.xml.rels':rels([('table1','table','../tables/table1.xml')])})
    matrix=worksheet(rows([cell('A1','Requirement ID'),cell('B1','Requirement'),cell('C1','Response'),cell('D1','Notes')],[cell('A2','REQ-001'),cell('B2','Unicode support'),cell('C2','待填写'),cell('D2','=literal response')],[cell('A3','REQ-002'),cell('B3','Offline operation'),cell('C3','Pending')]),'<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="C2:C3"><formula1>"Pending,Compliant,Not compliant"</formula1></dataValidation></dataValidations>')
    yield 'response-matrix',workbook([('Responses',matrix,'visible')],extras={'customXml/item1.xml':'<response-metadata preserve="exact">synthetic</response-metadata>'})
    budget=worksheet(rows([cell('A1','Line'),cell('B1','Quantity'),cell('C1','Unit cost'),cell('D1','Amount')],[cell('A2','People'),cell('B2',2),cell('C2',50),cell('D2',100,'$B2*C$2'),cell('F2',25,"'Rates'!B1*D4"),cell('G2',125,'AVERAGE(D2:D3)'),cell('H2',100,'MIN(D2:D3)'),cell('I2',150,'MAX(D2:D3)'),cell('J2',2,'COUNT(D2:D3)')],[cell('A3','Equipment'),cell('B3',3),cell('C3',50),cell('D3',150,'B3*C3')],[cell('A4','Total'),cell('D4',250,'SUM(D2:D3)')]))
    yield 'budget',workbook([('Budget',budget,'visible'),('Rates',worksheet(rows([cell('B1',0.1)])),'veryHidden')])
    diagnostics=worksheet(rows([cell('A1',10,'IF(1=1,10,0)'),cell('B1',15,'SUM(A1:A2)'),cell('C1',None,'UNKNOWN_LOCAL(A2)'),cell('D1',None,'D2+1'),cell('E1',raw='9007199254740993'),cell('F1',True),cell('G1',999,'SUM(A2:A2)')],[cell('A2',5),cell('B2',None,'1/0',error='#DIV/0!'),cell('C2',None,'SUM(C1,C3)'),cell('D2',None,'D1+1'),cell('E2','000123'),cell('F2','東京 / 中文'),cell('G2',None,'A2+1')],[cell('C3',5),cell('E3','=SUM(A1:A2)'),cell('F3')]))
    yield 'calculation-states',workbook([('Diagnostics',diagnostics,'visible')])
    yield 'dates-1900',workbook([('Dates',worksheet(rows(*[[cell(f'A{i}',v,style=1)] for i,v in enumerate([59,60,61,1462],1)])),'visible')])
    yield 'dates-1904',workbook([('Dates',worksheet(rows([cell('A1',0,style=1),cell('B1',1,'A1+1',1)],[cell('A2',1,style=1)])),'visible')],date1904=True)
    chart='<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$3</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>'
    anchor='<xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
    drawing=f'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{R}"><xdr:twoCellAnchor>'+anchor+'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="chart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor><xdr:twoCellAnchor>'+anchor+'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Image"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="image"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>'
    data=worksheet(rows([cell('A1','Category'),cell('B1','Amount')],[cell('A2','People'),cell('B2',125),cell('D2','Pending')],[cell('A3','Equipment'),cell('B3',80)]),'<mergeCells count="1"><mergeCell ref="D4:E4"/></mergeCells><dataValidations count="1"><dataValidation type="whole" sqref="B2:B3"><formula1>0</formula1><formula2>1000000</formula2></dataValidation></dataValidations><drawing r:id="drawing"/>')
    yield 'preservation-rich',workbook([('Data',data,'visible'),('Hidden',rates,'hidden')],names={'BudgetAmount':'Data!$B$2'},extras={'xl/worksheets/_rels/sheet1.xml.rels':rels([('drawing','drawing','../drawings/drawing1.xml')]),'xl/drawings/drawing1.xml':drawing,'xl/drawings/_rels/drawing1.xml.rels':rels([('chart','chart','../charts/chart1.xml'),('image','image','../media/image1.png')]),'xl/charts/chart1.xml':chart,'xl/media/image1.png':base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aWQAAAABJRU5ErkJggg=='),'customXml/item1.xml':'<metadata>Preserve exact synthetic content</metadata>'})
    yield 'protected',workbook([('Protected',worksheet(rows([cell('A1',125)]),'<sheetProtection sheet="1" objects="1"/>'),'visible')],protected=True)
    for count in [10000,100000]:
      body=''.join(f'<row r="{r}">'+''.join(cell(f'{chr(65+c)}{r}',r+c) for c in range(9))+cell(f'J{r}',2*r+1,f'A{r}+B{r}')+'</row>' for r in range(1,count//10+1))
      yield f'scale-{count}',workbook([('Scale',worksheet(body),'visible')])
    body=''.join(f'<row r="{r}">'+cell(f'A{r}',r,None if r==1 else f'A{r-1}+1')+'</row>' for r in range(1,10001))
    yield 'dependency-chain',workbook([('Chain',worksheet(body),'visible')])
    yield 'opaque-media',workbook([('Data',worksheet(rows([cell('A1',125)])),'visible')],extras={'xl/media/opaque.bin':random.Random(20260922).randbytes(16*1024*1024)})

def build(directory):
    directory.mkdir(parents=True,exist_ok=True)
    identities=[]
    for name,parts in corpus():
      file=directory/(name+'.xlsx')
      with zipfile.ZipFile(file,'w',compression=zipfile.ZIP_DEFLATED,compresslevel=6) as archive:
        for member,value in sorted(parts.items()):
          info=zipfile.ZipInfo(member,(1980,1,1,0,0,0));info.compress_type=zipfile.ZIP_DEFLATED;info.external_attr=0o600<<16
          archive.writestr(info,value.encode() if isinstance(value,str) else value)
      identities.append({'id':name,'file':file.name,'bytes':file.stat().st_size,'sha256':hashlib.sha256(file.read_bytes()).hexdigest()})
    (directory/'generated.json').write_text(json.dumps({'generator':'stdlib OOXML; independent of workbook implementation','generatorSha256':hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),'fixtures':identities},indent=2)+'\n')
    return identities
if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--output',type=Path,required=True);args=parser.parse_args()
    print(json.dumps({'fixtures':build(args.output)}))
