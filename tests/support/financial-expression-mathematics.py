"""Independent actual binary64 operands + repeated-period financial recurrence; no vendor output."""
import decimal,json,struct
D=decimal.Decimal;decimal.getcontext().prec=160
bits=lambda n:struct.pack('>d',n).hex()
source={'B1':0.12,'B2':5.0,'B3':1000.0,'B4':30.0}
rate=source['B1']/12.0;nper=source['B2']*12.0
out=[]
for name,formula,values in [('PMT','PMT(B1/12,B2*12,B3)',[rate,nper,source['B3']]),('PV','PV(B1/12,B2*12,-B4)',[rate,nper,-source['B4']]),('FV','FV(B1/12,B2*12,-B4,-B3,1)',[rate,nper,-source['B4'],-source['B3'],1.0])]:
    args=list(map(D.from_float,values))+[D(0)]*(5-len(values));r,n,flow,balance,timing=args;g=D(1);a=D(0)
    for _ in range(int(n)):g*=1+r;a=a*(1+r)+1
    due=1+r*timing;v=-(flow*g+balance)/(due*a)if name=='PMT'else -(flow*due*a+balance)/g if name=='PV'else -(balance*g+flow*due*a)
    nearest=float(v);out.append({'name':name,'formula':formula,'operands':values,'operandBits':list(map(bits,values)),'independentRecurrence':str(v),'mathematicalNearest':nearest,'mathematicalNearestBits':bits(nearest),'actualExcelReference':None})
print(json.dumps({'schema':'workbook.xlsx.financial-expression-sample-mathematics.v1','sourceLiterals':source,'scope':'Each authored operand rounded in Python binary64 in expression order, then160digit Decimal.from_float period recurrence; no generated vendor reference','cases':out},indent=2))
