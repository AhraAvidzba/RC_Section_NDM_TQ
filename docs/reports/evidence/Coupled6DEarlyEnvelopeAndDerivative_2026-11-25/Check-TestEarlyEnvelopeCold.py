"""ДЛЯ ТЕСТОВ. Проверяет настоящий холодный исход745 и фактический выбор ранней H-параметризации, не объявляя несущую способность."""
from pathlib import Path
from runpy import run_path
from zipfile import ZipFile
from hashlib import sha256
import json,re,numpy as np

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold745EarlyGrowingEnvelope_20261125';parent=base/'Cold743UnverifiedCumulativeRecovery_20261125'
helpers=run_path(str(here/'Check-TestPersistentHEvidence.py'));read=run_path(str(here/'Read-TestColdRawCapture.py'))
regression,text=helpers['TestNative'](root,'WorkingGuideRegressions')
assert regression==dict(passed=105,failed=0,macros=3,timeout600=False,own_process_closed=True)
native,text=helpers['TestNative'](root,'HBackendStock')
compiled_record=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled_record['WorkbookSHA256']
if native['timeout600']:
    outcome=dict(native=native,returned_workflow_result_missing=True,results_reopen_unverified=True)
else:
    saved,=re.findall(r'(?m)^RESULTS_SAVE_REOPEN: .*?equal=True; file=([^\r\n]+)',text)
    (root/'SavedWorkbook.xlsm').write_bytes(Path(saved).read_bytes())
    with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:compiled=book.read('xl/vbaProject.bin')
    outcome=helpers['TestStock'](repo,root,compiled)
frames,raw=read['TestHistory'](root/'Native/Stock.RawHistory.bin')
old,_=read['TestHistory'](parent/'Native/Stock.RawHistory.bin')
assert len(frames[0]['fields'])==len(old[0]['fields'])==4
for key in ('fields','fractions','lambdas'):assert np.array_equal(frames[0][key],old[0][key]),key
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledNewtonSolver.cls',
             'tests/modTestCoupledElasticState.bas'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows'](root/'Native/Stock.Checkpoints.csv')
details=[';'.join(row[1:]) for row in rows if row[0]=='DETAILS']
assert details and details[0].startswith('Мера деформации: материальная координата H огибающих;')
frame=frames[-1]
proof=dict(regression=regression,cold745=outcome,raw_history=raw,raw_frames=len(frames),accepted_states=len(frame['fields']),
    all_four_initial_force_states_exact743=True,initial_H_control_selected_before_first_controlled_probe=True,
    last_accepted_lambda=float(frame['lambdas'][-1]),maximum_accepted_lambda=float(max(frame['lambdas'])),
    final_control_details=details[-1],source_workbook_unchanged=True,
    independent_tensor_reintegration_pending_at_cold_recording=True,material_Newton_test_module_exact743=True,
    no_physical_capacity_claimed=True,no_working_promotion=True,whole_goal_completed=False)
(root/'EarlyEnvelopeColdProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('745',len(frame['fields']),'ACCEPTED STATES; INITIAL H SELECTED; MAX',proof['maximum_accepted_lambda'],
      'SECONDS',outcome.get('seconds'),'CAPACITY',outcome.get('capacity'),'NO PHYSICAL LIMIT CLAIM')
