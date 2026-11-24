"""ДЛЯ ТЕСТОВ. Сравнивает собранную касательную с полным откликом одной неизменной истории на двух шагах разностей."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold738TailFullDifferences_20261125';parent=base/'Cold737TailMatrixPathRetry_20261125'
proof=json.loads((parent/'TailMatrixProof.json').read_bytes())
assert proof['all70_tensor_output_bitidentical_to_verified735'] and proof['matrices'][-1]['numerical_rank_SVD']==17
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'FullDifferenceCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json','CompilePointFailure.json','SavedCompilationRecovery.json'):
    if (root/name).exists():(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Private Sub TestWriteFrozenTailMatrices\(.*?^End Sub',code)
assert method.count('    For mode = 0 To 2\n')==1
changed=method.replace('    For mode = 0 To 2\n','    For mode = 2 To 2\n',1)
anchor='        mReverseSolverTest.EndNewtonLineTest\n';assert changed.count(anchor)==1
changed=changed.replace(anchor,'        If ready Then TestWriteTailDifferenceSamples equations, zero, residual\n'+anchor,1)
code=code.replace(method,changed,1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает все столбцы производной полного отклика на двух шагах; каждый образец сразу закрывается для сохранения частичного свидетельства.
Private Sub TestWriteTailDifferenceSamples(ByVal equations As ICoupledEquationSystem, ByRef zero() As Double, ByRef assembledResidual() As Double)
    Dim file As String: file = mHardPreparedPathTest & ".TailDifferences.bin"
    Dim handle As Integer: handle = FreeFile
    Open file For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMTAILDIFF00001"
    Dim size As Long: size = UBound(zero)
    Dim stepCount As Long: stepCount = 2
    Dim sampleCount As Long: sampleCount = 1 + stepCount * (2 * size + 1)
    Put #handle, , magic: Put #handle, , size: Put #handle, , sampleCount
    TestWriteReverseBinaryVector handle, assembledResidual
    Close #handle
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim sample() As Double, residual() As Double, unused() As Double, ready As Boolean
    sample = zero
    ready = equations.TryEvaluate(sample, False, residual, unused)
    TestAppendTailDifference file, 0#, 0, 0, ready, residual
    Dim stepIndex As Long, column As Long, sign As Long, stepSize As Double
    For stepIndex = 1 To stepCount
        stepSize = 0.000001: If stepIndex = 2 Then stepSize = 0.0001
        For column = 1 To size
            For sign = -1 To 1 Step 2
                sample = zero: sample(column) = sign * stepSize
                ready = equations.TryEvaluate(sample, False, residual, unused)
                TestAppendTailDifference file, stepSize, column, sign, ready, residual
                If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Разностная проба изменила принятую историю."
            Next sign
        Next column
        sample = zero
        ready = equations.TryEvaluate(sample, False, residual, unused)
        TestAppendTailDifference file, stepSize, 0, 0, ready, residual
    Next stepIndex
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Дописывает исходные Double одной пробы; отсутствие материального отклика сохраняется флагом, а не подменяется нулевой производной.
Private Sub TestAppendTailDifference(ByVal file As String, ByVal stepSize As Double, ByVal column As Long, ByVal sign As Long, _
        ByVal ready As Boolean, ByRef residual() As Double)
    Dim handle As Integer: handle = FreeFile
    Open file For Binary Access Write As #handle
    Seek #handle, LOF(handle) + 1
    Dim flag As Long: flag = Abs(CLng(ready))
    Put #handle, , stepSize: Put #handle, , column: Put #handle, , sign: Put #handle, , flag
    If ready Then TestWriteReverseBinaryVector handle, residual
    Close #handle
End Sub
'''
assert max(map(len,code.splitlines()))<1024
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Update-IsolatedOptimizerCode.ps1';compiler=(root/name).read_text('utf-8-sig')
anchor='''    # ДЛЯ ТЕСТОВ: целевая книга закрывается при существующей другой книге только в этом собственном процессе Excel.
    $taskCloseAnchor=$excel.Workbooks.Add()
    $taskCloseAnchor.Saved=$true
    Write-Output 'COMPILE_OWN_EMPTY_CLOSE_ANCHOR_CREATED'
''';assert compiler.count(anchor)==1
compiler=compiler.replace(anchor,'',1)
target='    # ДЛЯ ТЕСТОВ: проект уже скомпилирован; ссылки его объектов освобождаются до закрытия книги-владельца.\n'
assert compiler.count(target)==1
compiler=compiler.replace(target,anchor+'\n'+target,1)
values[name]=compiler.replace('\n','\r\n').encode('utf-8-sig')
name='Run-TestPostCrackRawReplay.ps1';runner=(root/name).read_text('utf-8-sig');assert runner.count('-TimeoutSeconds 240')==1
values[name]=runner.replace('-TimeoutSeconds 240','-TimeoutSeconds 600',1).replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,b in values.items():(root/name).write_bytes(b);prep['inputs'][name]=sha256(b).hexdigest().upper()
prep['directed_scope'].update(case='FrozenForwardFullResidualDifferences',actual_import_this_pass=['modTestCoupledElasticState'],
    native_numeric_kernel_exact737=True,source_states=70,source_secant_ratio=1.25,
    normalized_difference_steps=[1e-6,1e-4],all61_columns_both_signs=True,expected_full_residual_samples=247,
    samples_flushed_individually=True,accepted_revision_unchanged_required=True,
    compiler_close_anchor_created_before_VBE_release=True,no_new_capacity_search=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('738 FULL RESIDUAL DIFFERENCES STAGED;247 SAMPLES; SAME70 HISTORY ANDNUMERICAL CORE')
