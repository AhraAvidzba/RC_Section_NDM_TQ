"""ДЛЯ ТЕСТОВ. Готовит частную выдачу из действующей пользовательской книги с заменой ровно четырёх рабочих классов ASD."""
from pathlib import Path
from hashlib import sha256
import json, shutil

repo = Path.cwd()
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1067WorkingASDNoObserver_20261127'
root = base / 'Cold1071WorkingASDOutput_20261127'
assert not root.exists()
proof = json.loads((parent / 'WorkingASDNoObserverProof.json').read_bytes())
observer = json.loads((base / 'Cold1068WorkingASDObserverMutation_20261127/WorkingASDObserverMutationProof.json').read_bytes())
assert proof['all6_no_observer_results_bitexact_verified_reference']
assert observer['all6_mutating_observer_results_bitexact_verified_reference']
book = (repo / 'workbook/output/RC_Section_NDM.xlsm').read_bytes()
dump = (repo / 'workbook/output/VBA_All_Code.txt').read_bytes()
assert sha256(book).hexdigest().upper() == 'E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
for directory in ('tests', 'src/Solver', 'Native', 'Temp', 'docs/regression', 'workbook/output', 'Before'):
    (root / directory).mkdir(parents=True, exist_ok=True)
shutil.copytree(parent / 'tools', root / 'tools')
classes = ['CCoupledDeformationContinuation', 'CCoupledLoadContinuation', 'CCoupledNewtonSolver', 'CCoupledSectionAnalysis']
for name in classes:
    path = f'src/Solver/{name}.cls'
    (root / path).write_bytes((repo / path).read_bytes())
(root / 'docs/regression/Diagnostic.xlsm').write_bytes(book)
(root / 'workbook/output/RC_Section_NDM.xlsm').write_bytes(book)
(root / 'Before/RC_Section_NDM.xlsm').write_bytes(book)
(root / 'Before/VBA_All_Code.txt').write_bytes(dump)
shutil.copyfile(parent / 'Guard-IsolatedOffice.ps1', root / 'Guard-IsolatedOffice.ps1')
compiler = (parent / 'Compile-TestWorkingASD.ps1').read_text('utf-8-sig')
(root / 'Compile-TestWorkingASDOutput.ps1').write_bytes(compiler.replace('\n', '\r\n').encode('utf-8-sig'))
(root / 'ParentCodeUpdate.json').write_text(json.dumps(dict(WorkbookSHA256=sha256(book).hexdigest().upper()), indent=2) + '\n', encoding='utf-8')
# Полный дамп снимается с сохранённой и повторно открытой книги, пока собственный процесс компилятора ещё открыт.
updater = root / 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1'
code = updater.read_text('utf-8-sig')
anchor = "    Write-Output 'COMPILE_AFTER_CLOSE'\n"
assert code.count(anchor) == 1
reopen = '''
    $book=$excel.Workbooks.Open($path,0,$true)
    if (!$book.ReadOnly) {throw 'Output verification requires a read-only reopen.'}
    $writer=[IO.StreamWriter]::new((Join-Path (Split-Path -Parent $path) 'VBA_All_Code.txt'),$false,[Text.UTF8Encoding]::new($true))
    try {
        $writer.WriteLine('VBA PROJECT EXPORT');$writer.WriteLine('Workbook: RC_Section_NDM.xlsm');$writer.WriteLine('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
        $project=$book.VBProject;$components=$project.VBComponents
        foreach($component in $components) {
            $module=$component.CodeModule
            $writer.WriteLine('='*100);$writer.WriteLine('COMPONENT: '+$component.Name);$writer.WriteLine('TYPE: '+$component.Type);$writer.WriteLine('LINES: '+$module.CountOfLines);$writer.WriteLine('='*100)
            if($module.CountOfLines){$writer.WriteLine($module.Lines(1,$module.CountOfLines))}
            $writer.WriteLine()
        }
        $excel.VBE.ActiveVBProject=$project
        $compile=$excel.VBE.CommandBars.FindControl(1,578)
        if ($compile.Enabled) {throw 'Saved and reopened output is not compiled.'}
        [ordered]@{SavedReopened=$true;ReadOnly=$book.ReadOnly;Components=$components.Count;Compiled=$true;VbaDumpFromActualReopenedBook=$true} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path (Split-Path -Parent $path) 'OutputReopenProof.json') -Encoding UTF8
    } finally {$writer.Dispose()}
    foreach($taskVbeObject in @($compile,$module,$component,$components,$project)) {
        if($null -ne $taskVbeObject -and [Runtime.InteropServices.Marshal]::IsComObject($taskVbeObject)){
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskVbeObject)
        }
    }
    $taskVbeObject=$null;$compile=$null;$module=$null;$component=$null;$components=$null;$project=$null
    $book.Close($false);$book=$null
    Write-Output 'OUTPUT_READONLY_REOPEN_AND_FULL_CODE_EXPORT_VERIFIED'
'''
code = code.replace(anchor, anchor + reopen, 1)
updater.write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
shutil.copyfile(Path(__file__), root / Path(__file__).name)
prep = dict(parent=str(parent), source_public_book_sha256=sha256(book).hexdigest().upper(),
    source_public_dump_sha256=sha256(dump).hexdigest().upper(),
    only_four_working_classes_replaced=classes, no_private_test_module_imported=True,
    inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
            for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix() != 'docs/regression/Diagnostic.xlsm'})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('PUBLIC-BASED ASD OUTPUT PREPARED;FOUR WORKING CLASSES;USER DATA RETAINED;NO TEST INSTRUMENTATION')
