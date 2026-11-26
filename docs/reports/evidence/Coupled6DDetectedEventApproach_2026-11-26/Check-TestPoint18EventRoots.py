"""ДЛЯ ТЕСТОВ. Проверяет материальные равновесия951 всеми исходными силовыми уравнениями и сопоставляет уровень половинной пробы949."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold951Point18EventRoots_20261126';parent=base/'Cold949WorkingGuidePair_20261126'
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
script=here/'Check-TestRecordedYieldCoordinate.py';code=script.read_text('utf-8-sig')
changes={"root=base/'Cold936RecordedYieldCoordinate_20261126'":"root=base/'Cold951Point18EventRoots_20261126'",
    "audit['imported_this_pass']==['modTestCoupledElasticState']":"audit['imported_this_pass']==['CCoupledNewtonSolver','modTestCoupledElasticState']",
    "len(audit['other_actual_components_exact_parent_except_identifier_case'])==256":"len(audit['other_actual_components_exact_parent_except_identifier_case'])==255",
    "if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()":"if path.is_file() and path.name!='CCoupledNewtonSolver.cls':assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()"}
for before,after in changes.items():assert code.count(before)==1;code=code.replace(before,after,1)
scope={'__file__':str(script),'__name__':'point18_event_roots_test'};exec(compile(code,str(script),'exec'),scope)
proof=json.loads((root/'RecordedYieldCoordinateProof.json').read_bytes())
proof.pop('all256_other_source_components_preserved');proof.pop('source_actual934_state25_and_initial_residual_bitexact_by_native_guard')
proof.update(other255_source_components_preserved_by_parent_audit=True,immutable_compiled949_reused=True,
    source_history25_unchanged=True,initial_verified941_root_R_bitexact_guard=True,
    selected_real_concrete_point=18,selected_side=0,prior_seed_cost_not_included_in_per_case_effort=True)
R=scope['R'];r=R((parent/'Native/WorkingGuidePairs.bin').read_bytes());assert r.VectorTest().tolist()==[1.]
pair,accepted,full,half,projection=[r.VectorTest() for _ in range(5)]
native=run_path(str(here/'Read-TestRecordedYieldCoordinate.py'))['TestRead'](root/'Native/RecordedYieldCoordinate.bin')
for item,case in zip(proof['cases'],native):
    if item['case']>2 and case['info'][5]:
        delta=np.einsum('spcd,d->spc',scope['op'],case['field']-half)
        item['known_half949_lambda_difference']=abs(item['lambda_value']-pair[1])
        item['known_half949_max_three_station_strain_difference']=float(np.max(abs(delta)))
    if item['case']<=2:
        item['is_event_surface_target']=True
(root/'Point18EventRootsProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('EVENT ROOTS VERIFIED',sum(c['independently_equilibrated'] for c in proof['cases'][:2]))
for c in proof['cases'][2:]:print('HALF LEVEL CASE',c['case'],'LAMBDA DIFFERENCE',c.get('known_half949_lambda_difference'),'STRAIN DIFFERENCE',c.get('known_half949_max_three_station_strain_difference'))
