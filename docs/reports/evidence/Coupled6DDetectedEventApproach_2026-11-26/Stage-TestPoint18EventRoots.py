"""ДЛЯ ТЕСТОВ. Ищет силовое равновесие на найденной поверхности18 и на уровне действительной половинной пробы949."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold949WorkingGuidePair_20261126';root=base/'Cold951Point18EventRoots_20261126'
source=base/'Cold941OriginalControlReturn_20261126'
proof=json.loads((source/'OriginalControlReturnProof.json').read_bytes());assert proof['original_equilibria_verified']==2
chosen=run_path(str(here/'Read-TestOriginalControlReturn.py'))['TestRead'](source/'Native/OriginalControlReturn.bin')[0]
assert chosen['info'][1] and chosen['info'][2]
cross=json.loads((base/'Cold950GuidePairBaseScan_20261126/AcceptedYieldCrossingProof.json').read_bytes())
half=next(p for p in cross['crossings'] if p['endpoint']=='half');assert half['point']==18 and half['side']==0 and half['after_trial_yield_MPa']>0
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/RecordedYieldSeed.bin').read_bytes());r.position=0
meta,current,accepted,residual,projection=[r.VectorTest() for _ in range(5)];assert r.position==len(r.raw)
meta[0]=half['point'];meta[1]=half['side'];meta[2]=0.;meta[9]=half['after_trial_yield_MPa'];meta[3]=chosen['info'][3]
current=chosen['field'];residual=chosen['values'][:60].copy();residual[56:]/=meta[8]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
updates={'Input/RecordedYieldSeed.bin':b''.join(vector(v) for v in (meta,current,accepted,residual,projection)),
    'SourceOriginalRoot.bin':(source/'Native/OriginalControlReturn.bin').read_bytes(),
    'SourceAcceptedYieldCrossingProof.json':json.dumps(cross,sort_keys=True,indent=2).encode(),
    Path(__file__).name:Path(__file__).read_bytes(),'Read-TestOriginalControlReturn.py':(here/'Read-TestOriginalControlReturn.py').read_bytes()}
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunWorkingGuidePair';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestSolveRecordedYieldCoordinate',1).replace('\n','\r\n').encode('utf-8-sig')
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='Point18ActualLoadingEventAndKnownHalfLevel',immutable_compiled949=True,no_new_compilation=True,
    point_selected_from_all129_source_scan=True,point=18,side=0,material_targets_MPa=[0.,float(meta[9])],
    initial_field_is_verified941_full_root=True,initial_native_R_bitexact_guard=True,
    half_level_from_actual949_state_not_arbitrary=True,source_history25_and_six_force_path_unchanged=True,
    original30_steps_per_case=True,no_earlier_effort_hidden_in_per_case_counts=True,
    no_material_history_acceptance=True,physical_capacity_certified=False,no_fresh_cold_claim=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'POINT18 EVENT F0 AND ACTUAL HALF LEVEL',meta[9],'SAME MATERIAL/HISTORY/SIX FORCE PATH')
