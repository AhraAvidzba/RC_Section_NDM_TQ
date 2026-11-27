"""ДЛЯ ТЕСТОВ: повторяет тот же QR-расчёт после устранения внешней блокировки файла наблюдателем."""
from pathlib import Path
from hashlib import sha256
import json
import shutil

repo=Path.cwd();base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1160ASDSteelLinearTrustMy_20261127';root=base/'Cold1161ASDSteelLinearRetry_20261127'
assert not root.exists()
closed=json.loads((parent/'Case03/Native/RealColdReplay.txt.ProcessClosure.json').read_text('utf-8-sig'))
assert closed['ChildHasExited'] and closed['ChildExitCode']==1
assert 'CCoupledMaterialProgressTest VBAProject Permission denied' in (parent/'Case03/Native/RealColdReplay.txt').read_text('utf-8-sig')
record=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
book=(parent/'docs/regression/Diagnostic.xlsm').read_bytes();assert sha256(book).hexdigest().upper()==record['WorkbookSHA256']
for folder in ['src','tests','Before','Native','Temp','Input','docs/regression','workbook/output']:
    (root/folder).mkdir(parents=True,exist_ok=True)
shutil.copytree(parent/'tools',root/'tools')
shutil.copyfile(parent/'Guard-IsolatedOffice.ps1',root/'Guard-IsolatedOffice.ps1')
for relative in ['Before/RC_Section_NDM.xlsm','docs/regression/Diagnostic.xlsm','workbook/output/RC_Section_NDM.xlsm']:
    (root/relative).write_bytes(book)
for name in ['Case01','Case02','Case03','Case04']:
    for folder in ['Input','Native','docs/regression']:(root/name/folder).mkdir(parents=True,exist_ok=True)
    shutil.copytree(parent/name/'Input',root/name/'Input',dirs_exist_ok=True)
    for relative in ['Case.json','Native/MainCapacity.Operators.csv']:
        shutil.copyfile(parent/name/relative,root/name/relative)
source=base/'Cold1157ASDSingleStepTangentMy_20261127/tests/CCoupledMaterialProgressTest.cls'
code=source.read_text('utf-8-sig');before='    Close #handle: handle = 0\n    Exit Sub'
assert code.count(before)==1
code=code.replace(before,'''    Close #handle: handle = 0
    FileCopy mPath, mPath & ".Accepted." & Format$(mCalls, "000000") & ".bin" ' Копия читается после закрытия собственного файла; внешнего читателя и конкуренции за доступ нет.
    Exit Sub''',1)
(root/'tests/CCoupledMaterialProgressTest.cls').write_text(code,encoding='utf-8-sig',newline='\r\n')
compiler=(parent/'Compile-TestASDSteelLinearTrustMy.ps1').read_text('utf-8-sig');before="@('CCoupledNewtonSolver')"
assert compiler.count(before)==1;compiler=compiler.replace(before,"@('CCoupledMaterialProgressTest')",1)
(root/'Compile-TestASDSteelLinearRetry.ps1').write_text(compiler,encoding='utf-8-sig',newline='\r\n')
shutil.copyfile(parent/'Run-TestASDSteelLinearTrustMy.ps1',root/'Run-TestASDSteelLinearRetry.ps1')
(root/'ParentCodeUpdate.json').write_text(json.dumps({'WorkbookSHA256':record['WorkbookSHA256']})+'\n',encoding='utf-8')
pins={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'}
(root/'Preparation.json').write_text(json.dumps({'inputs':pins,'base_git_commit':'751116a','expected_components':263,
    'imported_components':['CCoupledMaterialProgressTest'],'all_numerical_classes_exact_1160':True,
    'only_change':'Test-only FileCopy of accepted history after closing own file; no external reader',
    'cold_from_zero':True,'materials_limits_tolerances_and_probe_budget_unchanged':True,'public_book_unchanged':True},sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('SAME COLD QR CALCULATION PREPARED;ONLY TEST OBSERVER CHANGED;NO EXTERNAL READER')
