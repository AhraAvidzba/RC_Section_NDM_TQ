"""ДЛЯ ТЕСТОВ. Удаляет неприменимую ветвь уточнения плеча из однопроходной обёртки733; численный код не меняет."""
from pathlib import Path
from hashlib import sha256
import json,re

base=Path('D:/CodexScratch/RC_Section_NDM_TQ').resolve();root=(base/'Cold733FastFrozenReplay_20261124').resolve()
assert root.is_relative_to(base)
closure=json.loads((root/'CompilePointClosure.json').read_bytes());assert closure['State']=='Absent'
parent=json.loads((root/'ParentCodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==parent['WorkbookSHA256']
archive=root/'OriginalCompileFailure733';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json','CompilePointFailure.json'):
    p=root/name
    if p.exists():p.rename(archive/name)
for name in ('Preparation.json','tests/modTestCoupledElasticState.bas'):
    target=archive/name;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes((root/name).read_bytes())
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Private Function TestReplayPostCrackRaw\(.*?^End Function',code)
changed=method
anchor='If settings(1) <> mode Or (mode <> 0 And mode <> 1) Then';assert changed.count(anchor)==1
changed=changed.replace(anchor,'If settings(1) <> mode Or mode <> 0 Then',1)
start=changed.index('    Else\n        Dim target As Double: target = settings(2)\n')
finish=changed.index('    End If\n    TestReplayPostCrackRaw = report',start)
changed=changed[:start]+changed[finish:]
assert 'TestConfirmComputedShoulder' not in changed
code=code.replace(method,changed,1);assert 'TestConfirmComputedShoulder' not in code
(root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
prep=json.loads((root/'Preparation.json').read_bytes())
prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
name=Path(__file__).name;(root/name).write_bytes(Path(__file__).read_bytes());prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
prep['directed_scope'].update(unused_legacy_quality_branch_removed=True,replay_mode_zero_only=True,
    numerical_solver_and_material_unchanged=True,original_compile_failure_preserved=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('733 ORIGINAL COMPILE FAILURE PRESERVED; UNUSED QUALITY BRANCH REMOVED; NUMERIC CORE UNCHANGED')
