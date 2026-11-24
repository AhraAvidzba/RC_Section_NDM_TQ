"""ДЛЯ ТЕСТОВ. Проверяет исход холодного743, фактическое продолжение и сохранённую книгу без подмены численной остановки пределом."""
from pathlib import Path
from runpy import run_path
from zipfile import ZipFile
from hashlib import sha256
import json,re,numpy as np

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold743UnverifiedCumulativeRecovery_20261125';parent=base/'Cold741OrdinaryFirstComplex_20261125'
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
frame=frames[-1];reference=old[-1]
# Уточнение подтверждённого плеча вправе заменить хвост; до первой новой ветви исходные12 обязаны воспроизвестись.
initial=next(f for f in frames if len(f['fields'])>=len(reference['fields']))
for key in ('fields','fractions','lambdas'):
    assert np.array_equal(initial[key][:len(reference['fields'])],reference[key]),key
count=0
for i in range(min(len(frame['fields']),len(reference['fields']))):
    if not all(np.array_equal(frame[key][i],reference[key][i]) for key in ('fields','fractions','lambdas')):break
    count+=1
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledNewtonSolver.cls',
             'tests/modTestCoupledElasticState.bas'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows'](root/'Native/Stock.Checkpoints.csv')
details=[';'.join(row[1:]) for row in rows if row[0]=='DETAILS']
unverified=[int(v) for line in details for v in re.findall(r'неподтверждённых спадов: (\d+)',line)]
proof=dict(regression=regression,cold743=outcome,raw_history=raw,raw_frames=len(frames),
    accepted_states=len(frame['fields']),exact_source_prefix_states=count,
    initial_before_recovery_all12_source_states_exact=True,
    last_accepted_lambda=float(frame['lambdas'][-1]),maximum_accepted_lambda=float(max(frame['lambdas'])),
    max_reported_unverified_declines=max(unverified,default=0),
    independent_tensor_reintegration_pending=True,no_physical_capacity_claimed=True,
    material_Newton_and_test_module_exact741=True,no_working_promotion=True,whole_goal_completed=False)
proof['compiled_source_workbook_unchanged']=True
(root/'CumulativeRecoveryColdProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('743',len(frame['fields']),'ACCEPTED STATES; EXACT PREFIX',count,'MAX',proof['maximum_accepted_lambda'],
      'UNVERIFIED DECLINES',proof['max_reported_unverified_declines'],'NO PHYSICAL LIMIT CLAIM')
