"""ДЛЯ ТЕСТОВ. Сопоставляет повторное чтение материала950 с сохранённым состоянием25 и фактической парой949."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold950GuidePairBaseScan_20261126';parent=base/'Cold949WorkingGuidePair_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
script=here/'Check-TestWorkingGuidePair.py';code=script.read_text('utf-8-sig')
# Переиспользуются только определения сборки и метрики из уже проверенного949, без новой нативной проверки949.
cut="r=R((root/'Native/WorkingGuidePairs.bin').read_bytes())";assert code.count(cut)==1
scope={'__file__':str(script),'__name__':'guide_pair_base_independent_test'};exec(compile(code[:code.index(cut)],str(script),'exec'),scope)
inner=scope['inner'];R=scope['R'];op=scope['op'];source=scope['source'];L=inner['L']
rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows']
r=R((root/'Input/ArcScanFields.bin').read_bytes());r.position=0;assert r.VectorTest().tolist()==[3.]
fields=[r.VectorTest() for _ in range(3)];assert r.position==len(r.raw)
cases=json.loads((root/'Input/ArcScanCases.json').read_bytes());np.testing.assert_array_equal(fields[0],source['fields'][-1])
stress={}
for row in rows(root/'Native/ArcScan.csv'):
    if row[0]=='S' and float(row[2])==1:
        assert int(row[5])==0;key=int(row[1]),int(row[4]),int(row[3]);assert key not in stress;stress[key]=np.array(row[6:12],float)
saved={}
for row in rows(root/'Native/EnvelopeEnergy.Accepted.csv'):
    if row[0]=='S' and int(row[1])==25:saved[int(row[2]),int(row[3])]=np.array(row[4:10],float)
items=[]
for case,field in zip(cases,fields):
    sigma=np.array([[stress[case['case'],s,p] for p in range(1,44)] for s in (-1,0,1)])
    a=inner['TestAssemble'](field,case['lambda_value'],sigma)
    item=dict(name=case['name'],lambda_value=case['lambda_value'],physical_residual_MPa=a['physical_norm'],six_force_residual_MPa=a['six_force_norm'])
    if case['case']==1:
        difference=[dict(side=s,point=p,max_stress_difference_MPa=float(max(abs(stress[1,s,p]-saved[s,p])))) for s in (-1,0,1) for p in range(1,44)]
        item['maximum_stored_vs_reread_stress_difference_MPa']=max(d['max_stress_difference_MPa'] for d in difference)
        item['largest_stress_differences']=sorted(difference,key=lambda d:d['max_stress_difference_MPa'],reverse=True)[:5]
    items.append(item)
pair=json.loads((parent/'WorkingGuidePairProof.json').read_bytes())['pairs'][0]
for i in (1,2):assert abs(items[i]['physical_residual_MPa']-pair['states'][i-1]['independent_total_residual_MPa'])<2e-11
proof=dict(native=native,immutable_compiled949_reused=True,readonly_no_new_NDS_or_acceptance=True,
    original25_history_unchanged=True,states=items,physical_capacity_certified=False)
(root/'GuidePairBaseScanProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for p in items:print(p['name'],'R',p['physical_residual_MPa'],'FORCE',p['six_force_residual_MPa'],'STORED DIFF',p.get('maximum_stored_vs_reread_stress_difference_MPa'))
