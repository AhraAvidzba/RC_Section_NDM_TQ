"""ДЛЯ ТЕСТОВ. Вызывает существующий поиск события от фактической неудачной итерации934 без будущего равновесия и без компиляции нового кода."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
import json,numpy as np

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold960ResolvedFieldProgress_20261126';root=base/'Cold961FailedPredictorEvent_20261126'
origin=base/'Cold935Point11Matrices_20261126'
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=R((origin/'Input/ActualRejectedNewton.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
matrix=np.array([r.VectorTest() for _ in range(61)]);ending=r.VectorTest();count=int(r.VectorTest()[0])
probes=[[r.VectorTest() for _ in range(3)] for _ in range(count)];assert r.position==len(r.raw)
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
source=read(parent/'Input/Reverse1.bin');old=read(origin/'Input/Reverse1.bin')
for name in ('calibration','options','offset','direction','fields','lambdas'):np.testing.assert_array_equal(source[name],old[name])
np.testing.assert_array_equal(accepted,source['fields'][-1]);assert len(source['fields'])==25 and meta[2]==1.1024974179997635
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
context=[source['lambdas'][-1],meta[2],source['options'][3],source['options'][4],source['options'][2]]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector'];projection=matrix[-1,:60]/meta[4]
updates={'Input/DetectedEvent.bin':vector(context)+vector(current)+vector(projection),
    'SourceActualRejectedNewton.bin':(origin/'Input/ActualRejectedNewton.bin').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
anchor='modTestCoupledElasticState.TestRunFieldProgressSteps';assert runner.count(anchor)==1
updates[name]=runner.replace(anchor,'modTestCoupledElasticState.TestRunDetectedLoadingBoundary',1).replace('\n','\r\n').encode('utf-8-sig')
input_origin=dict(origin=str(origin),source_file='Input/ActualRejectedNewton.bin',source_sha256=sha256((origin/'Input/ActualRejectedNewton.bin').read_bytes()).hexdigest().upper(),
    exact_original25_geometry_materials_force_path_and_fields=True,source_meta=meta.tolist(),source_residual_inf_MPa=float(max(abs(residual))),
    starting_state_is_failed_Newton_iterate_not_an_equilibrium=True,no_future_equilibrium_in_event_input=True,
    input_context=context,predictor_field=current.tolist(),projection=projection.tolist())
updates['FailedPredictorInputOrigin.json']=(json.dumps(input_origin,sort_keys=True,indent=2)+'\n').encode()
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExistingEventSearchFromActualFailedNewtonIterate934',actual_import_this_pass=[],actual_component_count=257,
    all257_actual_components_and_saved_book_bitexact960=True,existing_TestRunDetectedLoadingBoundary_reused=True,
    no_new_numerical_code_or_compilation=True,original_force_tolerances_and30_steps=True,
    no_future_equilibrium_in_consumed_event_input=True,source_history25_bitexact935=True,
    no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
statepath=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statepath.read_bytes());assert not state['active_jobs']
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    previous_goal_turn_classification='progress:958-960 and105 focused regression assertions verified/published93a2054;8 incoming field steps vs1 with original lambda growth criterion, original force/steel/half checks retained and lambda uncertainty preserved; still supplied guide and no physical capacity. Next authoritative inspection identified existing Arc pre-event recovery and internal1e-4 accuracy scaling.',
    next_action='961 prepared from bitexact compiled960: existing generic event API receives only actual failed934 iterate and original25 history, no supplied future equilibrium. Run/verify event outcome, then existing OperationalArc path/call budget/accuracy separation from same history. Goal active.')
raw=(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n').encode('utf-8')
for path in (statepath,here/'Continuation666_Resume20261123.json'):path.write_bytes(raw)
print(root,'ACTUAL FAILED934 PREDICTOR;NO FUTURE ROOT;ALL257 ACTUAL SOURCES AND BOOK REUSED BITEXACT')
