"""ДЛЯ ТЕСТОВ. Фиксирует реальный холодный шестикомпонентный контроль740 без подмены тайм-аута пределом."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold740CurrentComplexStock_20261125';parent=base/'Cold739PseudoTimeProbe_20261125'
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'HBackendStock')
assert native['timeout600'] and native['own_process_closed'] and 'TIMING_STATES:' not in text
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Native/Stock.RawHistory.bin')
assert raw['incomplete_tail_bytes']==0;frame=frames[-1]
code=(root/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
values,=re.findall(r'TestHardShearLoads = Array\(([^\n]+)\)',code)
forces=[float(v.strip().replace('#','')) for v in values.split(',')]
assert len(forces)==6 and all(v!=0 for v in forces)
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls',
    'src/Materials/CCoupledConcreteResponse.cls','tests/modTestCoupledElasticState.bas'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
proof=dict(native=native,source_raw_history=raw,accepted_states=len(frame['fields']),raw_frames=len(frames),
    native_maximum_accepted_lambda=float(max(frame['lambdas'])),last_accepted_lambda=float(frame['lambdas'][-1]),
    source_internal_loads_N_Nmm=forces,all_six_components_nonzero=True,
    original_macro='RunTestStockHTrace -> RunTestCoupledWorkbookTiming(Both,Excel)',
    compiled_source_exact739=True,no_new_compilation_or_material_state_reuse=True,
    no_axial_history_or_external_seed_used=True,source_workbook_unchanged=True,
    returned_workflow_result_missing=True,capacity_limit_not_confirmed=True,
    results_reopen_not_verified=True,independent_tensor_reintegration_not_yet_done=True,
    not_a_clean_performance_benchmark_due_to_checkpoint_observer=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'CurrentComplexStopProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('740 REAL COLD ALL6 TIMEOUT600; ACCEPTED',proof['accepted_states'],'MAX',proof['native_maximum_accepted_lambda'],
    'NO RETURNED WORKFLOW LIMIT; REOPEN UNVERIFIED')
