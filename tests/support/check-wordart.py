"""Single owner fixture adapter for the independent WordArt graph oracle."""
import sys,json,tempfile,subprocess,pathlib
expected=[{'id':'11111111-1111-4111-8111-111111111111','sample':sys.argv[3],'text':'Edited "&é\nNext','altText':'Edited WordArt alternative text','anchor':{'kind':'oneCell','cell':'B4','dxEmu':100,'dyEmu':200,'widthEmu':3048000,'heightEmu':1143000}}]
with tempfile.NamedTemporaryFile(mode='w',suffix='.json') as f:
 json.dump(expected,f);f.flush();result=subprocess.run([sys.executable,str(pathlib.Path(__file__).with_name('check-wordart-controls.py')),sys.argv[1],sys.argv[2],f.name],capture_output=True,text=True);print(result.stdout,end='');print(result.stderr,end='',file=sys.stderr);sys.exit(result.returncode)
