"""ДЛЯ ТЕСТОВ. Сохраняет неудачную компиляцию и делает чтение диагностической панели VBE ограниченным и безопасным."""
from pathlib import Path
from hashlib import sha256
import json

root = Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold753IntegratedColdFrozenReplay_20261125')
parent = Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold752IntegratedGuideDirected_20261125')
archive = root / 'CompileAttempt1Failure'
closure_path = root / 'CompilePointClosure.json'
if not closure_path.exists():
    closure_path = archive / 'CompilePointClosure.json'
closure = json.loads(closure_path.read_bytes())
assert closure['State'] == 'Absent'
book = root / 'docs/regression/Diagnostic.xlsm'
assert book.read_bytes() == (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
failure_path = root / 'CompilePointFailure.json'
if not failure_path.exists():
    failure_path = archive / 'CompilePointFailure.json'
failure = json.loads(failure_path.read_text('utf-8-sig'))
assert 'line 190' in failure['ScriptStackTrace']
archive.mkdir(exist_ok=True)
for name in ('CompilePoint.log', 'CompilePointClosure.json', 'CompilePointFailure.json'):
    if (root / name).exists():
        assert not (archive / name).exists()
        (root / name).rename(archive / name)
    else:
        assert (archive / name).exists()
name = 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1'
path = root / name
raw = path.read_bytes()
if (archive / 'Update-IsolatedOptimizerCode.ps1').exists():
    assert (archive / 'Update-IsolatedOptimizerCode.ps1').read_bytes() == raw
else:
    (archive / 'Update-IsolatedOptimizerCode.ps1').write_bytes(raw)
code = raw.decode('utf-8-sig').replace('\r\n', '\n')
old = """    if ($compile.Enabled) {$compile.Execute()}
    if ($compile.Enabled) {
        $pane=$excel.VBE.ActiveCodePane
        [int]$firstLine=0;[int]$firstColumn=0;[int]$lastLine=0;[int]$lastColumn=0
        $pane.GetSelection([ref]$firstLine,[ref]$firstColumn,[ref]$lastLine,[ref]$lastColumn)
        Write-Output ('COMPILE_LOCATION: module='+$pane.CodeModule.Parent.Name+'; line='+$firstLine+'; column='+$firstColumn)
        Write-Output $pane.CodeModule.Lines([Math]::Max(1,$firstLine-2),5)
        throw 'Private VBA compilation failed.'
    }"""
new = """    if ($compile.Enabled) {$compile.Execute()}
    # ДЛЯ ТЕСТОВ: ожидание завершения команды ограничено; активная команда после него остаётся ошибкой сборки.
    for ($taskCompileProbe=0;$taskCompileProbe -lt 40 -and $compile.Enabled;$taskCompileProbe++) {
        Start-Sleep -Milliseconds 250
    }
    if ($compile.Enabled) {
        $pane=$excel.VBE.ActiveCodePane
        if ($null -ne $pane) {
            [int]$firstLine=0;[int]$firstColumn=0;[int]$lastLine=0;[int]$lastColumn=0
            $pane.GetSelection([ref]$firstLine,[ref]$firstColumn,[ref]$lastLine,[ref]$lastColumn)
            Write-Output ('COMPILE_LOCATION: module='+$pane.CodeModule.Parent.Name+'; line='+$firstLine+'; column='+$firstColumn)
            Write-Output $pane.CodeModule.Lines([Math]::Max(1,$firstLine-2),5)
        } else {Write-Output 'COMPILE_LOCATION_UNAVAILABLE: active code pane is null; compilation not accepted.'}
        throw 'Private VBA compilation failed or has not completed within the bounded wait.'
    }"""
assert code.count(old) == 1
changed = code.replace(old, new, 1)
path.write_bytes(changed.replace('\n', '\r\n').encode('utf-8-sig'))
prep = json.loads((root / 'Preparation.json').read_bytes())
prep['inputs'][name] = sha256(path.read_bytes()).hexdigest().upper()
prep['inputs'][Path(__file__).name] = sha256(Path(__file__).read_bytes()).hexdigest().upper()
(root / Path(__file__).name).write_bytes(Path(__file__).read_bytes())
prep['directed_scope']['compiler_diagnostic_bounded_wait_and_null_pane_guard'] = True
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
(root / 'CompileDiagnosticRepair.json').write_text(json.dumps(dict(first_failure=failure, first_closure=closure,
    original_book_unchanged=True, original_compilation_not_accepted=True,
    before_sha256=sha256(raw).hexdigest().upper(), after_sha256=sha256(path.read_bytes()).hexdigest().upper(),
    numerical_sources_changed=False), ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('753 FAILED ATTEMPT PRESERVED; ORIGINAL BOOK EXACT752; NO NUMERICAL CHANGES; BOUNDED COMPILER DIAGNOSTIC FIX')
