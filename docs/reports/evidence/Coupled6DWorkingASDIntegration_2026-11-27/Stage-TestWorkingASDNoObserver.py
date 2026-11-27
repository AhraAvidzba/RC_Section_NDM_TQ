"""ДЛЯ ТЕСТОВ. Собирает рабочие четыре класса ASD в отдельной книге и проверяет обычный прямой вызов без наблюдателей и файлового ввода-вывода ядра."""
from pathlib import Path
from hashlib import sha256
import json, shutil

repo = Path.cwd()
here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1066ASDPureNDirectColdRepeat_20261127'
root = base / 'Cold1067WorkingASDNoObserver_20261127'
assert not root.exists()
cold = json.loads((parent / 'ColdDirectRepeatProof.json').read_bytes())
assert cold['all6_cold_histories_bitexact'] and cold['all6_expected_direct_states_verified']
compiled = json.loads((parent / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
book = (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
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
classes = ['CCoupledDeformationContinuation', 'CCoupledLoadContinuation', 'CCoupledNewtonSolver', 'CCoupledSectionAnalysis']
for name in classes:
    path = f'src/Solver/{name}.cls'
    data = (repo / path).read_bytes()
    assert b'modTest' not in data and b'TestCapture' not in data
    (root / path).write_bytes(data)
compiler_source = base / 'Cold1065ASDPureNDirectEarlyPrefix_20261127/Compile-TestDirectRecoveryPrefix.ps1'
compiler = compiler_source.read_text('utf-8-sig')
before = "-ModuleNames 'CCoupledLoadContinuation'"
assert compiler.count(before) == 1
replacement = '-ModuleNames @(' + ','.join("'" + name + "'" for name in classes) + ')'
compiler = compiler.replace(before, replacement, 1)
(root / 'Compile-TestWorkingASD.ps1').write_bytes(compiler.replace('\n', '\r\n').encode('utf-8-sig'))
(root / 'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent),
    variant='working-source-no-observer',actual_import_this_pass=classes), indent=2) + '\n', encoding='utf-8')
shutil.copyfile(Path(__file__), root / Path(__file__).name)
promotion = base / 'CleanASDWorkingPromotion_20261127/PromotionPreparation.json'
shutil.copyfile(promotion, root / 'WorkingPromotionPreparation.json')
prep = dict(parent=str(parent), directed_scope=dict(case='WorkingASDDirectNoObserver',
    actual_import_this_pass=classes, same_six_inputs_materials_tolerances=True,
    default_observer_is_Nothing=True, test_module_unchanged=True,
    kernel_has_no_test_exports=True, physical_acceptance_not_assumed=True, no_CDP_run=True),
    inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
            for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix() != 'docs/regression/Diagnostic.xlsm'})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('WORKING ASD NO-OBSERVER TEST PREPARED;EXACT4 WORKTREE CLASSES;UNCHANGED6 CASES;NATIVE ACCEPTANCE PENDING')
