"""ДЛЯ ТЕСТОВ. Проверяет рабочий необязательный наблюдатель ASD существующим классом, намеренно изменяющим только полученные копии."""
from pathlib import Path
from hashlib import sha256
import json, shutil

repo = Path.cwd()
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1067WorkingASDNoObserver_20261127'
root = base / 'Cold1068WorkingASDObserverMutation_20261127'
assert not root.exists()
proof = json.loads((parent / 'WorkingASDNoObserverProof.json').read_bytes())
assert proof['all6_no_observer_results_bitexact_verified_reference'] and proof['native']['own_process_closed']
book = (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
compiled = json.loads((parent / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256(book).hexdigest().upper() == compiled['WorkbookSHA256']
for directory in ('src', 'tests', 'tools', 'Input', 'workbook'):
    shutil.copytree(parent / directory, root / directory)
for directory in ('Native', 'Temp', 'docs/regression'):
    (root / directory).mkdir(parents=True, exist_ok=True)
for index in range(1, 7):
    case = root / f'Case{index:02d}'
    shutil.copytree(parent / f'Case{index:02d}/Input', case / 'Input')
    for directory in ('Native', 'docs/regression'):
        (case / directory).mkdir(parents=True, exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1', 'Run-TestASDDirectPureN.ps1'):
    shutil.copyfile(parent / name, root / name)
(root / 'docs/regression/Diagnostic.xlsm').write_bytes(book)
shutil.copyfile(parent / 'docs/regression/CodeUpdate.json', root / 'ParentCodeUpdate.json')
module_path = 'tests/modTestSelectableMaterialCore.bas'
code = (parent / module_path).read_text('utf-8-sig')
before = '    Set meta = analysis.SolveTarget(loads, ChrW$(&H3BB) & "*N", spec, field, resultants, True, state, trace)\n'
after = '''    Dim observer As New CCoupledProgressObserverTest, progress As ICoupledContinuationProgress
    observer.TestInitialize "", True: Set progress = observer
    Set meta = analysis.SolveTarget(loads, ChrW$(&H3BB) & "*N", spec, field, resultants, True, state, trace, progress)
    If observer.TestCallCount < 1 Then Err.Raise 5, , "Рабочий наблюдатель не получил ни одного снимка."
    Dim observerHandle As Integer: observerHandle = FreeFile
    Open folder & "\\Native\\ObserverCalls.bin" For Binary Access Write As #observerHandle
    Dim observerMagic As String * 16: observerMagic = "NDMOBSCALLS00001": Put #observerHandle, , observerMagic
    TestWriteVector observerHandle, Array(observer.TestCallCount, 1#)
    Close #observerHandle
'''
assert code.count(before) == 1
code = code.replace(before, after, 1)
(root / module_path).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
compiler = (parent / 'Compile-TestWorkingASD.ps1').read_text('utf-8-sig')
before = "@('CCoupledDeformationContinuation','CCoupledLoadContinuation','CCoupledNewtonSolver','CCoupledSectionAnalysis')"
assert compiler.count(before) == 1
compiler = compiler.replace(before, "'modTestSelectableMaterialCore'", 1)
(root / 'Compile-TestWorkingASDObserver.ps1').write_bytes(compiler.replace('\n', '\r\n').encode('utf-8-sig'))
(root / 'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent), variant='mutating-observer',
    numerical_classes_reused_bitexact=True, actual_import=['modTestSelectableMaterialCore']), indent=2) + '\n', encoding='utf-8')
(root / 'ObserverMutationMode.json').write_text(json.dumps(dict(existing_test_class='CCoupledProgressObserverTest',
    mutation_of_fields_trace_and_meta=True, numerical_classes_unchanged=True), indent=2) + '\n', encoding='utf-8')
shutil.copyfile(Path(__file__), root / Path(__file__).name)
prep = dict(parent=str(parent), directed_scope=dict(case='WorkingASDObserverMutation',
    actual_import_this_pass=['modTestSelectableMaterialCore'],
    existing_observer_test_class_reused=True, original_cases_materials_and_tolerances_unchanged=True,
    no_CDP_run=True, observer_isolation_not_assumed=True),
    inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
            for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix() != 'docs/regression/Diagnostic.xlsm'})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('WORKING ASD MUTATING OBSERVER TEST PREPARED;EXISTING TEST CLASS;NO NUMERICAL CHANGES')
