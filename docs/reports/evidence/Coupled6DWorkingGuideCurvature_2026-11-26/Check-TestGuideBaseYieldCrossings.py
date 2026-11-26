"""ДЛЯ ТЕСТОВ. Ищет смены знака поверхности между исходным состоянием25 и фактической полной/половинной парой950."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold950GuidePairBaseScan_20261126';rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows'](root/'Native/ArcScan.csv')
values={(int(r[1]),float(r[2]),int(r[3]),int(r[4])):float(r[6]) for r in rows if r[0]=='Y' and r[5]=='True'}
crossings=[]
for case in (2,3):
    for point in range(1,44):
        for side in (-1,0,1):
            left=values[1,1.,point,side];right=values[case,1.,point,side]
            if left < -3.2e-11 and right > 3.2e-11:
                crossings.append(dict(endpoint=('full' if case==2 else 'half'),point=point,side=side,
                    real_concrete=point<=25,before_trial_yield_MPa=left,after_trial_yield_MPa=right))
assert len(crossings)==2 and all(p['point']==18 and p['side']==0 for p in crossings)
proof=dict(crossings=crossings,source_native_arc_scan_sha256=sha256((root/'Native/ArcScan.csv').read_bytes()).hexdigest().upper(),
    threshold_MPa=3.2e-11,all129_locations_scanned=True,diagnosis='Base-to-endpoint loading change; full and half endpoints share the same active signatures.',
    intermediate_equilibrium_on_surface_not_yet_solved=True,not_a_physical_failure_or_capacity_claim=True)
(root/'AcceptedYieldCrossingProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(json.dumps(proof,ensure_ascii=False,indent=2))
