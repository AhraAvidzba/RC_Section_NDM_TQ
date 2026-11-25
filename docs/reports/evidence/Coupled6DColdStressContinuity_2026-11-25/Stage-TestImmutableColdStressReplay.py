"""ДЛЯ ТЕСТОВ. Переиспользует точную скомпилированную книгу с одним только читающим продолжением после экспорта напряжений."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from argparse import ArgumentParser
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--source',required=True);p.add_argument('--root',required=True);p.add_argument('--proof',required=True)
a=p.parse_args();source=base/a.source;root=base/a.root;compiled=base/'Cold855BalancedInitialControl_20261125'
proof=json.loads((source/a.proof).read_bytes());cold=proof['cold']
assert cold.get('own_process_closed',False) or cold.get('native',{}).get('own_process_closed',False)
partial=bool(cold.get('workflow_result_missing',False))
if partial:assert cold['timeout_seconds']==600
code=(compiled/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub',code)
assert method[method.index('    mPtcSuppressTest = True\n'):]=='''    mPtcSuppressTest = True
    TestExportFastFrozenReplay
    mPtcSuppressTest = False
    TestReadInitialControlCases
    mPolishExecutedTest = True
End Sub'''
readonly,=re.findall(r'(?ms)^Private Sub TestReadInitialControlCases\(.*?^End Sub',code)
assert 'subject.TestInitialControlSelection' in readonly and 'If before(i) <> after(i) Then Err.Raise' in readonly
assert 'If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise' in readonly
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](source/'Native/Stock.RawHistory.bin')
assert raw['incomplete_tail_bytes']==0 and len(frames[-1]['fields'])==proof['accepted_states']
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
original=read(source/'Input/Reverse1.bin');compiled_input=read(compiled/'Input/Reverse1.bin')
import numpy as np
for key in ('calibration','options','offset','direction'):np.testing.assert_array_equal(original[key],compiled_input[key])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,compiled,1e-6,1,1e-6,0)
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(compiled/'docs/regression/Diagnostic.xlsm').read_bytes()
prep=json.loads((root/'Preparation.json').read_bytes());values={
 'Input/Reverse1.bin':run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,frames[-1]),
 'Input/SourceColdRawHistory.bin':(source/'Native/Stock.RawHistory.bin').read_bytes(),
 'SourceColdProof.json':(source/a.proof).read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
compilation=json.loads((compiled/'docs/regression/CodeUpdate.json').read_bytes())
prep['directed_scope']=dict(case='VerifiedColdStressOnlyReplay',actual_import_this_pass=compilation['Modules'],
 actual_component_count=256,source_root=source.name,source_proof=a.proof,source_states=len(frames[-1]['fields']),
 source_workflow_incomplete=partial,source_history_exact=True,immutable_compiled_book_reused=True,
 compiled_source_root=compiled.name,actual_imports_this_replay=[],
 original_compilation_imports_preserved_for_runner_guard=True,
 only_readonly_initial_selector_cases_after_frozen_export=True,no_new_NDS_or_capacity_search=True,
 controller_Trace_not_called=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'EXACT',len(frames[-1]['fields']),'COLD STATES;SAME COMPILED BOOK;NO NEW IMPORT/COMPILE/SEARCH;PARTIAL',partial)
