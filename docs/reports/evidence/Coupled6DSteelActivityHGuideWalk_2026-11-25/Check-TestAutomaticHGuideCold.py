"""ДЛЯ ТЕСТОВ. Проверяет настоящий холодный исход и сохранённую книгу, отдельно отмечая применимость прежнего независимого интегрирования."""
from pathlib import Path
from runpy import run_path
from zipfile import ZipFile
import json,re,numpy as np

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold847AutomaticHGuideCold_20261125';prior=base/'Cold832GatedLoadingCold_20261125';replay=base/'Cold835GatedColdReplay_20261125'
helpers=run_path(str(here/'Check-TestPersistentHEvidence.py'));native,text=helpers['TestNative'](root,'HBackendStock')
assert native['failed']==0 and not native['timeout600'] and native['own_process_closed']
saved,=re.findall(r'(?m)^RESULTS_SAVE_REOPEN: .*?equal=True; file=([^\r\n]+)',text)
saved=Path(saved).resolve();assert saved.is_relative_to(root.resolve()) and saved.suffix=='.xlsm'
(root/'SavedWorkbook.xlsm').write_bytes(saved.read_bytes())
with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:compiled=book.read('xl/vbaProject.bin')
outcome=helpers['TestStock'](repo,root,compiled);assert outcome['all_saved_sources_exact']==256
read=run_path(str(here/'Read-TestColdRawCapture.py'));frames,raw=read['TestHistory'](root/'Native/Stock.RawHistory.bin')
oldframes,_=read['TestHistory'](prior/'Native/Stock.RawHistory.bin');frame=frames[-1];old=oldframes[-1]
assert raw['incomplete_tail_bytes']==0
identical=all(np.array_equal(frame[k],old[k]) for k in ('fields','fractions','lambdas'))
proof=dict(native=native,cold=outcome,raw_history=raw,accepted_states=len(frame['fields']),
    maximum_accepted_lambda=float(max(frame['lambdas'])),last_accepted_lambda=float(frame['lambdas'][-1]),
    cold832_fields_fractions_lambdas_bit_identical=identical,independent_cold_stress_replay_pending=not identical,
    warm_fields_not_supplied_to_stock_solver=True,physical_capacity_certified=False,whole_goal_completed=False)
if identical:
    previous=json.loads((replay/'OperationalColdReplayProof.json').read_bytes())
    assert previous['source_cold_exact'] and previous['accepted_states']==len(frame['fields'])
    for name in ('src/Solver/CCoupledElasticState.cls','src/Materials/CCoupledConcreteReturn.cls'):
        assert (root/name).read_bytes()==(replay/name).read_bytes()
    def TestMethods(path):
        """Сопоставляет все существовавшие методы материального отклика без изменения арифметики."""
        rows=re.findall(r'(?ms)^((?:Private|Friend|Public) (?:Function|Sub|Property Get) (\w+)\b.*?^End (?:Sub|Function|Property))',path.read_text('utf-8-sig'))
        assert len(rows)==len({n for _,n in rows});return {n:s for s,n in rows}
    for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls'):
        a,b=TestMethods(replay/name),TestMethods(root/name)
        assert all(a[k]==b[k] for k in a) and set(b)-set(a)=={'TryReadElasticPrincipalBoundary'}
    proof['independent_mechanics_reused_by_exact_fields_and_material_methods']=previous
if outcome['capacity']=='NumFail':
    assert not outcome['limit'] and abs(outcome['lower']-max(frame['lambdas']))<1e-11
    proof['NumFail_not_reclassified_as_failure']=True
(root/'AutomaticHGuideColdProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('ACTUAL COLD:',outcome['capacity'],'STATES',len(frame['fields']),'MAX',proof['maximum_accepted_lambda'],
      'SECONDS',outcome['seconds'],'SAVED/REOPENED256;EXACT PRIOR HISTORY',identical)
