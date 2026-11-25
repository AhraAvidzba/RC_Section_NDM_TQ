"""ДЛЯ ТЕСТОВ. Проверяет завершённый холодный штатный запуск, сохранение Results и фактический код повторно открытой книги."""
from pathlib import Path
from runpy import run_path
from zipfile import ZipFile
from argparse import ArgumentParser
import json
import re
import numpy as np

here=Path(__file__).resolve().parent; repo=Path.cwd()
p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True);a=p.parse_args();root=a.root
helpers=run_path(str(here/'Check-TestPersistentHEvidence.py'))
native,text=helpers['TestNative'](root,'HBackendStock')
if native['timeout600']:
    outcome=dict(native=native,returned_workflow_result_missing=True,physical_capacity_certified=False)
else:
    saved,=re.findall(r'(?m)^RESULTS_SAVE_REOPEN: .*?equal=True; file=([^\r\n]+)',text)
    saved=Path(saved).resolve();assert saved.is_relative_to(root.resolve()) and saved.suffix=='.xlsm'
    (root/'SavedWorkbook.xlsm').write_bytes(saved.read_bytes())
    with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:binary=book.read('xl/vbaProject.bin')
    outcome=helpers['TestStock'](repo,root,binary)
    assert outcome['all_saved_sources_exact']==256
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Native/Stock.RawHistory.bin')
assert raw['incomplete_tail_bytes']==0
frame=frames[-1]
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
prefix=0
for i in range(min(len(frame['fields']),len(source['fields']))):
    if not all(np.array_equal(frame[k][i],source[k][i]) for k in ('fields','fractions','lambdas')):break
    prefix+=1
rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows'](root/'Native/Stock.Checkpoints.csv')
details=[';'.join(x[1:]) for x in rows if x[0]=='DETAILS']
latest=details[-1]
match=re.search(r'вызовов проверенной знаковой дуги (\d+); принятых направлений (\d+)',latest)
assert match
proof=dict(cold=outcome,raw_history=raw,raw_frames=len(frames),accepted_states=len(frame['fields']),
    maximum_accepted_lambda=float(max(frame['lambdas'])),last_accepted_lambda=float(frame['lambdas'][-1]),
    exact_prefix_with_warm_input=prefix,arc_calls=int(match[1]),arc_recovered_directions=int(match[2]),latest_details=latest,
    actual_cold_workflow_not_resumed_from_Input_Reverse=True,independent_cold_stress_replay_pending=True,
    physical_capacity_certified=False,whole_goal_completed=False)
if not native['timeout600']:
    assert outcome['capacity']=='NumFail' and not outcome['limit']
    assert abs(outcome['lower']-max(frame['lambdas']))<1e-11
(root/'OperationalColdProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('COLD',outcome.get('capacity','TIMEOUT'),'STATES',len(frame['fields']),'MAX',proof['maximum_accepted_lambda'],'ARC_CALLS',match[1],'RECOVERED',match[2],'SAVED256',outcome.get('all_saved_sources_exact'))
