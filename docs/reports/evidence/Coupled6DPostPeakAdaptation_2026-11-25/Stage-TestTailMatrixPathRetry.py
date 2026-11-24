"""ДЛЯ ТЕСТОВ. Повторяет736 после исправления только диагностического пути; исходный Runtime Error сохраняется."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold737TailMatrixPathRetry_20261125';parent=base/'Cold736FrozenTailMatrix_20261125'
text=(parent/'Native/PostCrackReplay.txt').read_text('utf-8-sig')
closure=json.loads((parent/'Native/PostCrackReplay.txt.ProcessClosure.json').read_bytes())
assert 'RUNTIME ERROR: VBAProject Path not found' in text and 'SOURCE_UNCHANGED: True' in text
assert closure['ChildHasExited'] and closure['ChildExitCode']==1 and closure['Passed']
assert closure['Closure']['FinalState']['State']=='Absent'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PathRetryCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='        Open mPolishRootTest & "\\Native\\TailMatrix" & CStr(mode) & ".bin" For Binary Access Write As #handle'
assert code.count(anchor)==1
code=code.replace(anchor,'        Open mHardPreparedPathTest & ".TailMatrix" & CStr(mode) & ".bin" For Binary Access Write As #handle',1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values['OriginalNativeFailure736.txt']=(parent/'Native/PostCrackReplay.txt').read_bytes()
values['OriginalNativeFailure736Closure.json']=(parent/'Native/PostCrackReplay.txt.ProcessClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,b in values.items():(root/name).write_bytes(b);prep['inputs'][name]=sha256(b).hexdigest().upper()
prep['directed_scope'].update(case='FrozenTailMatrixPathRetry',actual_import_this_pass=['modTestCoupledElasticState'],
    only_change_to736='diagnostic output path uses actual replay prefix instead of unset unrelated context',
    source_runtime_failure_preserved=True,numerical_core_and_all_inputs_exact736=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('737 ONLY DIAGNOSTIC OUTPUT PATH CORRECTED;736 FAILURE PRESERVED')
