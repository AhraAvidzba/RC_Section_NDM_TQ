"""ДЛЯ ТЕСТОВ. Дополняет архив полной проверкой71 полей и остановкой холодного поиска по96 пробам."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
out=repo/'docs/reports/evidence/Coupled6DShoulderCost_2026-11-24'
pins=json.loads((out/'Files.json').read_bytes())
assert all(sha256((out/n).read_bytes()).hexdigest().upper()==v for n,v in pins.items())
source=base/'Cold732ProvidedJacobianCapacity_20261124';replay=base/'Cold733FastFrozenReplay_20261124'
stop=json.loads((source/'ProvidedCapacityStopProof.json').read_bytes());verified=json.loads((replay/'FastFrozenReplayProof.json').read_bytes())
assert verified['original_binary_input_and_exported_fields71_match'] and verified['independent_fields']['states']==71
stop.update(independent_all71_fields_pending=False,independent_all71_fields_verified_via733=True,
    original_full_export_still_incomplete=True,analytical_force_misses=len(verified['analytical_force_misses']),
    maximum_analytical_force_error_N=max(p['error_N'] for p in verified['analytical_forces']))
(source/'ProvidedCapacityStopProof.json').write_text(json.dumps(stop,indent=2,sort_keys=True)+'\n',encoding='utf-8')
for number,root in (('732',source),('733',replay)):
    names=['Preparation.json','CompilePointClosure.json','CompiledSourceReuse.json','PointTraceSourceComposition.json',
        'docs/regression/CodeUpdate.json','Run.ndjson','ProvidedCapacityStopProof.json','FastFrozenReplayProof.json',
        'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls',
        'src/Materials/CCoupledConcreteResponse.cls','tests/modTestCoupledElasticState.bas']
    paths=[root/n for n in names if (root/n).is_file()]
    paths += [p for p in (root/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    if number=='733':
        paths += [root/'Input/PostCrackReplay.bin',root/'Input/PostCrackReplayOptions.bin']
        paths += [p for p in (root/'OriginalCompileFailure733').rglob('*') if p.is_file()]
        paths += [root/'Stage-TestFastFrozenReplay.py',root/'Repair-TestFastFrozenReplay.py']
    for p in paths:
        target=out/number/p.relative_to(root);assert not target.exists()
        target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(p.read_bytes())
helpers=['Check-TestProvidedCapacityStop.py','Stage-TestFastFrozenReplay.py','Repair-TestFastFrozenReplay.py',
    'Check-TestFastFrozenReplay.py',Path(__file__).name]
for name in helpers:
    target=out/name;assert not target.exists();target.write_bytes((here/name).read_bytes())
(out/'ColdStopAndReplay.json').write_text(json.dumps(dict(automatic732=stop,
    replay733_native=verified['native'],all71_states_independently_verified=True,
    growth_diagnosis_scope='reconstruction from consecutive final fields, same native spatial metric and load-resolution formula',
    growth_diagnosis_tail=verified['consecutive_final_history_growth_diagnosis'][-14:],
    physical_capacity_certified=False,whole_goal_completed=False),indent=2,sort_keys=True)+'\n',encoding='utf-8')
for p in out.rglob('*'):
    if p.is_file() and p.name!='Files.json':pins[str(p.relative_to(out)).replace('\\','/')]=sha256(p.read_bytes()).hexdigest().upper()
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
    automatic732_stop=stop,frozen733_replay=dict(native=verified['native'],all71_input_and_export_fields_match=True,
        all71_independent_weak_gauge_and_six_force_checks=True,analytical_misses=len(verified['analytical_force_misses']),
        tail66_69_field_prediction_passes_but_load_prediction_fails=True,
        tail70_71_both_growth_checks_pass=True,physical_capacity_certified=False),
    previous_goal_turn_classification='progress: cold732 now returns numerical96probe budget stop before export, correct refinedlocalpeak1.441395945659 accuracy.00096493 andknownerror4.11e-7 butno steel event; export later timedout600.733 validates all71 exactfields and independent mechanics in single8s nativeStressOnly replay;36analytical misses retained. Taildiagnosis shows load prediction curvature gate, not field prediction, delaysgrowth at66-69; bothpass70-71 onlyasbudgetends. Full6 goal preserved.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('732-733 APPENDED; ALL71 VERIFIED; NUMERICAL BUDGET STOP NOT PHYSICAL LIMIT; TOTAL ARCHIVE',len(pins))
