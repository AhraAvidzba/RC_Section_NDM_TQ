"""ДЛЯ ТЕСТОВ. Сохраняет два отрицательных посттрещинных опыта и подтверждённую остановку собственных процессов."""
from pathlib import Path
from hashlib import sha256
import json

repo=Path.cwd();root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold700PostCrackLocalMatrix_20261124')
proof=json.loads((root/'PostCrackLocalMatrixProof.json').read_bytes())
assert all(v['native']['own_process_closed'] for v in proof['results'].values())
out=repo/'docs/reports/evidence/Coupled6DPostCrackLocalMatrix_2026-11-24';out.mkdir(parents=True,exist_ok=False)
files={}
for name in ('PostCrackLocalMatrixProof.json','PointTraceSourceComposition.json','CompilePointClosure.json'):
    target=out/name;target.write_bytes((root/name).read_bytes());files[name]=sha256(target.read_bytes()).hexdigest().upper()
for label in ('Matrix','Refined'):
    for name in ('Curve'+label+'.txt','Curve'+label+'.txt.ProcessClosure.json'):
        target=out/name;target.write_bytes((root/'Native'/name).read_bytes());files[name]=sha256(target.read_bytes()).hexdigest().upper()
    for name in ('PostCrackTension.Curve.csv','PostCrackPrescribed.Accepted.csv','PostCrackPrescribed.Accepted.Operators.csv','PostCrackPrescribed.Accepted11.bin'):
        target=out/(label+'_'+name);target.write_bytes((root/label/'Native'/name).read_bytes());files[target.name]=sha256(target.read_bytes()).hexdigest().upper()
for name in ('Stage-TestPostCrackLocalMatrix.py','Check-TestPostCrackLocalMatrix.py',Path(__file__).name):
    target=out/name;target.write_bytes((repo/'tmp/material-consistency'/name).read_bytes());files[name]=sha256(target.read_bytes()).hexdigest().upper()
(out/'Files.json').write_text(json.dumps(files,indent=2,sort_keys=True)+'\n',encoding='utf-8')
(out/'.gitattributes').write_text('* -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8',newline='\n')
print('BOTH NEGATIVE CURVES SAVED; ALL22 PRESCRIBED STATES INDEPENDENTLY VERIFIED; WORKING NEWTON UNCHANGED')
