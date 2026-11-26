"""ДЛЯ ТЕСТОВ. Уточняет непрерывность отклика непосредственно вокруг первой поверхности фактического отказа54."""
from pathlib import Path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;file=here/'Stage-TestRecordedTailStressScan.py'
source=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold926RecordedTailStressScan_20261126/RecordedTailKinksProof.json')
proof=json.loads(source.read_bytes());crossing=proof['first_trial_surface_crossing']
assert crossing['point']==5 and crossing['side']==0
center=(crossing['fraction_lower']+crossing['fraction_upper'])/2
extras=[center+s*10.**(-k) for k in range(4,11) for s in (-1,1)]
code=file.read_text('utf-8-sig').replace('Cold926RecordedTailStressScan_20261126','Cold928RecordedTailBoundaryScan_20261126')
anchor='predictions=[accepted if f is None else current+f*delta[:60] for f in fractions]'
assert code.count(anchor)==1
code=code.replace(anchor,'fractions+='+repr(extras)+'\n'+anchor)
scope={'__file__':str(file),'__name__':'recorded_tail_boundary_scan_test'}
exec(compile(code,str(file),'exec'),scope)
root=scope['root'];prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in ((Path(__file__).name,Path(__file__).read_bytes()),('SourceRecordedTailKinksProof.json',source.read_bytes())):
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ActualTailFirstSurfaceContinuityScan',endpoint_fractions=scope['fractions'],
    first_trial_surface_crossing_fraction=center,localization_not_material_failure=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('BOUNDARY ENDPOINTS',len(scope['fractions']),'AROUND',center)
