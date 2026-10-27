"""Independent per-period recurrence, using actual decoded binary64 inputs.
No Excel value is generated or claimed. Distinct from decimal-text recurrence.
"""
import decimal,json,struct,sys
decimal.getcontext().prec=160
D=decimal.Decimal
plan=json.load(open(sys.argv[1])); output=[]
for case in plan['cases']:
    numeric=[float(value) for value in case['arguments']]
    args=[D.from_float(value) for value in numeric]+[D(0)]*(5-len(numeric))
    rate,nper,flow,balance,timing=args
    growth=D(1);annuity=D(0)
    for _ in range(int(nper)):
        growth*=1+rate
        annuity=annuity*(1+rate)+1
    due=1+rate*timing
    result=(-(flow*growth+balance)/(due*annuity) if case['name']=='PMT' else
            -(flow*due*annuity+balance)/growth if case['name']=='PV' else
            -(balance*growth+flow*due*annuity))
    nearest=float(result)
    if nearest==0: nearest=0.0
    output.append({'id':case['id'],'formula':case['formula'],
      'inputBits':[struct.pack('>d',v).hex() for v in numeric],
      'binary64InputDecimalRecurrence':str(result),'nearestBinary64Bits':struct.pack('>d',nearest).hex(),
      'nearestBinary64Value':nearest,'exactRationalVector':case['mathematicalReference']['exactRationalVector'],
      'actualExcelReference':None})
json.dump({'schema':'workbook.xlsx.financial-binary64-input-mathematics.v1','method':'160-digit Decimal.from_float inputs, independent repeated-period recurrence','excelRuntimeEvidence':False,'cases':output},sys.stdout,indent=2);print()
