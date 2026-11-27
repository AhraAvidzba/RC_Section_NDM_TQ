"""ДЛЯ ТЕСТОВ. Добавляет точный проверенный набор зависимостей ASD в копию действующей книги, сохраняя исходную неудачную попытку и все пользовательские данные."""
from pathlib import Path
from hashlib import sha256
from argparse import ArgumentParser
import json, shutil

repo = Path.cwd()
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
previous = base / 'Cold1071WorkingASDOutput_20261127'
parser = ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, default=base / 'Cold1072WorkingASDDependenciesOutput_20261127')
parser.add_argument('--verify-reopen-compile', action='store_true')
args = parser.parse_args()
root = args.root
assert not root.exists()
closure = json.loads((previous / 'CompilePointClosure.json').read_text('utf-8-sig'))
assert closure['State'] == 'Absent'
audit = json.loads((base / 'CleanASDWorkingPromotion_20261127/PublicDependencyInspection.json').read_bytes())
book = (repo / 'workbook/output/RC_Section_NDM.xlsm').read_bytes()
assert sha256(book).hexdigest().upper() == audit['public_sha256']
assert (previous / 'Before/RC_Section_NDM.xlsm').read_bytes() == book
for directory in ('tests', 'src', 'Native', 'Temp', 'docs/regression', 'workbook/output', 'Before'):
    (root / directory).mkdir(parents=True, exist_ok=True)
shutil.copytree(previous / 'tools', root / 'tools')
if args.verify_reopen_compile:
    updater = root / 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1'
    code = updater.read_text('utf-8-sig')
    before = "        if ($compile.Enabled) {throw 'Saved and reopened output is not compiled.'}"
    after = """        $taskReopenCompileInitiallyEnabled=[bool]$compile.Enabled
        if ($compile.Enabled) {$compile.Execute()}
        for ($taskReopenProbe=0;$taskReopenProbe -lt 40 -and $compile.Enabled;$taskReopenProbe++) {
            Start-Sleep -Milliseconds 250
        }
        if ($compile.Enabled) {throw 'Actual compilation of saved and reopened output failed.'} """
    assert code.count(before) == 1
    code = code.replace(before, after, 1)
    before = 'SavedReopened=$true;ReadOnly=$book.ReadOnly;Components=$components.Count;Compiled=$true;'
    assert code.count(before) == 1
    code = code.replace(before, before + 'CompileCommandInitiallyEnabled=$taskReopenCompileInitiallyEnabled;ReopenedCompileExecutedIfEnabled=$true;', 1)
    updater.write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
shutil.copytree(previous / 'Before', root / 'Before', dirs_exist_ok=True)
(root / 'workbook/output/RC_Section_NDM.xlsm').write_bytes(book)
(root / 'docs/regression/Diagnostic.xlsm').write_bytes(book)
for file in ('Guard-IsolatedOffice.ps1', 'ParentCodeUpdate.json'):
    shutil.copyfile(previous / file, root / file)
imports = []
for item in audit['changed_or_added']:
    name, path = item['name'], item['source']
    if path is None:
        assert name == 'modTestSelectableMaterialCore'
        continue
    # Эти два отличия подробно сохранены аудитом: ANSI-комментарий Bairán и отображение конечной Double-константы VBE.
    assert item['private_exact_current_source'] or name in ('CCoupledElasticState', 'CCrackFormationCalculator')
    destination = root / path
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes((repo / path).read_bytes())
    imports.append(name)
compiler = (previous / 'Compile-TestWorkingASDOutput.ps1').read_text('utf-8-sig')
before = "@('CCoupledDeformationContinuation','CCoupledLoadContinuation','CCoupledNewtonSolver','CCoupledSectionAnalysis')"
assert compiler.count(before) == 1
compiler = compiler.replace(before, '@(' + ','.join("'" + name + "'" for name in imports) + ')', 1)
(root / 'Compile-TestWorkingASDDependenciesOutput.ps1').write_bytes(compiler.replace('\n', '\r\n').encode('utf-8-sig'))
shutil.copyfile(base / 'CleanASDWorkingPromotion_20261127/PublicDependencyInspection.json', root / 'PublicDependencyInspection.json')
shutil.copyfile(Path(__file__), root / Path(__file__).name)
prep = dict(parent=str(previous), source_public_book_sha256=audit['public_sha256'],
    source_public_dump_sha256=sha256((repo / 'workbook/output/VBA_All_Code.txt').read_bytes()).hexdigest().upper(),
    imports=imports, total_imports=len(imports), original_public_components=audit['public_count'],
    expected_new_components=audit['private_count'] - 1, private_diagnostic_module_excluded=True,
    no_sheets_settings_or_data_written=True,
    verify_actual_compilation_after_reopen=args.verify_reopen_compile,
    inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
            for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix() != 'docs/regression/Diagnostic.xlsm'})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('PUBLIC-BASED ASD DEPENDENCIES PREPARED;IMPORTS', len(imports), ';COMPONENTS', prep['expected_new_components'])
