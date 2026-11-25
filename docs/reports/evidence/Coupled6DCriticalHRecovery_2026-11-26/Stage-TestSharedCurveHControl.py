"""ДЛЯ ТЕСТОВ. Переиспользует одну совместную метрику дуги, возврата и H-поворота и проверяет H-корни от истинной истории20."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold867FixedMetricKernel_20261126';root=base/'Cold874SharedCurveHControl20_20261126';verified=base/'Cold869FixedMetricNextScan_20261126'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0,verified/'Input/Reverse1.bin')
archive=root/'SharedCurveCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}

def replace(code,old,new):
    """Требует точного совпадения единственного исходного участка перед заменой."""
    assert code.count(old)==1,(old[:90],code.count(old));return code.replace(old,new,1)

name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
code=replace(code,'    lambdaBound = GeomMax(YIELD_EVENT_HALF_ERROR * Abs(fullLambda - referenceLambda), 8# * relativeTolerance * (1# + Abs(referenceLambda)))',
'''    Dim curveScale As Double, jointError As Double, stepError As Double
    If Not TryElasticArcScale(direction, curveScale) Then GoTo CountHalf
    lambdaBound = CurveLambdaBound(length, tensorError, fullLambda, referenceLambda, lambdaError, curveScale, _
        YIELD_EVENT_HALF_ERROR, relativeTolerance, jointError, stepError)''')
code=replace(code,'    details = details & "; полный/половинный возврат: ошибка поля " & CStr(tensorError) & "; ошибка lambda " & CStr(lambdaError) & "; допуск lambda " & CStr(lambdaBound)',
'''    details = details & "; полный/половинный возврат: ошибка поля " & CStr(tensorError) & "; совместная ошибка " & CStr(jointError) & _
        "; ошибка lambda " & CStr(lambdaError) & "; допуск lambda " & CStr(lambdaBound)''')
old='''        If tensorError > YIELD_EVENT_HALF_ERROR Or lambdaError > GeomMax(YIELD_EVENT_HALF_ERROR * Abs(fullLambda - acceptedLambda), _
                8# * relativeTolerance * (1# + Abs(acceptedLambda))) Then GoTo Unsuccessful'''
new='''        Dim curveScale As Double, jointError As Double, stepError As Double, lambdaBound As Double
        If Not TryElasticArcScale(direction, curveScale) Then GoTo Unsuccessful
        lambdaBound = CurveLambdaBound(length, tensorError, fullLambda, acceptedLambda, lambdaError, curveScale, _
            YIELD_EVENT_HALF_ERROR, relativeTolerance, jointError, stepError)
        If tensorError > YIELD_EVENT_HALF_ERROR Or lambdaError > lambdaBound Then GoTo Unsuccessful'''
code=replace(code,old,new)
code=replace(code,'mEnvelopeDetails = mEnvelopeDetails & "; поворот подтверждён половинным H; ошибка поля " & CStr(tensorError) & ", lambda " & CStr(lambdaError)',
'mEnvelopeDetails = mEnvelopeDetails & "; поворот подтверждён половинным H; ошибка поля " & CStr(tensorError) & ", совместная ошибка " & CStr(jointError) & ", lambda " & CStr(lambdaError)')
code+='''

' Ограничивает кривизну полного и половинного состояния в постоянной совместной метрике поля и нагрузки.
' Возвращает допуск множителя; отдельный предел ошибки поля и физические проверки сохраняет вызывающий алгоритм.
' alpha — тензорная норма упругого отклика на направление LoadPath; noise — исходное разрешение lambda.
Friend Function CurveLambdaBound(ByVal length As Double, ByVal fieldError As Double, ByVal factor As Double, _
        ByVal reference As Double, ByVal lambdaError As Double, ByVal alpha As Double, ByVal errorLimit As Double, _
        ByVal relativeTolerance As Double, ByRef jointError As Double, ByRef stepError As Double) As Double
    jointError = -1#: stepError = -1#
    If length <= 0# Or alpha <= 0# Or errorLimit <= 0# Or relativeTolerance <= 0# Or fieldError < 0# Or lambdaError < 0# Then _
        Err.Raise 5, "CCoupledNewtonSolver.CurveLambdaBound", "Не заданы положительные масштабы и точность контроля половинного шага."
    Dim radius As Double, fieldDefect As Double, allowance As Double, noise As Double
    radius = Sqr(length * length + alpha * alpha * (factor - reference) ^ 2)
    fieldDefect = fieldError * length
    allowance = (errorLimit * radius) ^ 2 - fieldDefect * fieldDefect
    noise = 8# * relativeTolerance * (1# + Abs(reference))
    CurveLambdaBound = noise
    If allowance > 0# Then CurveLambdaBound = GeomMax(noise, Sqr(allowance) / alpha)
    jointError = Sqr(fieldDefect * fieldDefect + (alpha * lambdaError) ^ 2) / radius
    Dim resolvedLambdaDefect As Double: resolvedLambdaDefect = lambdaError
    If lambdaError <= noise Then resolvedLambdaDefect = 0#
    stepError = GeomMax(fieldError, Sqr(fieldDefect * fieldDefect + (alpha * resolvedLambdaDefect) ^ 2) / radius)
End Function
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig');start=code.index('Private Function PairLambdaBound(');end=code.index('\nEnd Function',start)+len('\nEnd Function')
code=code[:start]+'''Private Function PairLambdaBound(ByVal length As Double, ByVal fieldError As Double, ByVal factor As Double, _
        ByVal reference As Double, ByVal lambdaError As Double) As Double
    PairLambdaBound = mSolver.CurveLambdaBound(length, fieldError, factor, reference, lambdaError, mReferenceScale, _
        FIELD_ERROR_LIMIT, mRelative, mJointError, mStepError)
End Function'''+code[end:]
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
entries=[]
entries.append('''Public Function TestReadSharedCurveCases() As String
    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Проверяет общий численный метод на аналитических переходах через известный максимум; силовая приёмка этим не заменяется.
    Dim inputHandle As Integer: inputHandle = FreeFile
    Open mPolishRootTest & "\\Input\\SharedCurveCases.bin" For Binary Access Read As #inputHandle
    Dim header As Variant: header = TestReadPolishVector(inputHandle)
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\SharedCurveCases.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMSHAREDCURVE01"
    Put #handle, , magic: TestWriteReverseBinaryVector handle, header
    Dim i As Long, x As Variant, jointError As Double, stepError As Double, bound As Double
    For i = 1 To CLng(header(1))
        x = TestReadPolishVector(inputHandle)
        bound = mReverseSolverTest.CurveLambdaBound(x(1), x(2), x(3), x(4), x(5), x(6), x(7), x(8), jointError, stepError)
        TestWriteReverseBinaryVector handle, x
        TestWriteReverseBinaryVector handle, Array(bound, jointError, stepError)
    Next i
    If Seek(inputHandle) <> LOF(inputHandle) + 1 Then Err.Raise 5, , "Лишние данные аналитических проверок метрики."
    Close #inputHandle: Close #handle
    TestReadSharedCurveCases = "TOTAL_SHARED_CURVE_CASES_TEST: passed=1; failed=0; analytical metric only; physical capacity not asserted"
End Function''')
entries.append('''Public Function TestRunHCornerPairs() As String
    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Проверяет три независимые полные/половинные H-пары от истинной истории20; история не принимается и предел не объявляется.
    Dim accepted() As Double, previous() As Double, beforeH As Double, baseline As Double, resolution As Double, afterH As Double
    Dim fraction As Double, factor As Double, previousFactor As Double, count As Long: count = mReverseFieldsTest.Count
    mReverseSolverTest.CopyAccepted accepted
    mReverseTraceTest.ReadPoint count, fraction, factor: mReverseTraceTest.ReadPoint count - 1, fraction, previousFactor
    Dim absTol As Double, relTol As Double: absTol = mReverseOptionsTest(3) * 0.0001: relTol = mReverseOptionsTest(4) * 0.0001
    If Not mReverseSolverTest.TryEnvelopeControlState(mReverseBaseTest, mReverseDirectionTest, factor, absTol, relTol, baseline, resolution) Then Err.Raise 5, , "Нет исходного H."
    If Not mReverseSolverTest.TryAcceptedEnvelopeSegment(previous, beforeH, afterH) Then Err.Raise 5, , "Нет исходной принятой пары H."
    If afterH <> baseline Or afterH <= beforeH Then Err.Raise 5, , "Для прогноза требуется фактический рост H."
    Dim alpha As Double
    If Not mReverseSolverTest.TryElasticArcScale(mReverseDirectionTest, alpha) Then Err.Raise 5, , "Нет постоянного масштаба LoadPath."
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\HCornerPairs.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMHCORNERPAIR01"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(count, factor, previousFactor, baseline, beforeH, resolution, alpha, revision)
    Dim logHandle As Integer: logHandle = FreeFile
    Open mPolishRootTest & "\\Native\\HCornerPairs.txt" For Output As #logHandle
    Dim caseIndex As Long, ratio As Double, target As Double, seed() As Double, projection() As Double, field() As Double, actual() As Double
    Dim i As Long, role As Long, solved As Boolean, returned As Double, probes As Long, used As Long, budget As Long, energy As Double
    Dim full() As Double, half() As Double, lf As Double, lh As Double, sf As Boolean, sh As Boolean
    Dim records As New Collection, record As Collection
    For caseIndex = 1 To 3
        ratio = 0.25 ^ (caseIndex - 1): used = 0
        For role = 1 To 2
            Dim portion As Double: portion = 1#: If role = 2 Then portion = 0.5
            target = baseline + portion * ratio * (afterH - beforeH): seed = accepted
            For i = 1 To mReverseSolverTest.PhysicalDegreeCount: seed(i) = accepted(i) + portion * ratio * (accepted(i) - previous(i)): Next i
            budget = 30 - used: solved = False: probes = 0: returned = 0#
            Dim iterations As Long: iterations = 0
            If budget > 0 Then
                solved = mReverseSolverTest.TryEnvelopeStep(mReverseBaseTest, mReverseDirectionTest, previous, factor, seed, _
                    factor + portion * ratio * (factor - previousFactor), target, returned, projection, absTol, relTol, budget, probes, False)
                iterations = mReverseSolverTest.IterationCount
            End If
            used = used + iterations: energy = 0#
            TestWriteReverseBinaryVector handle, Array(caseIndex, role, ratio, portion, target, Abs(CLng(solved)), returned, probes, iterations, used, revision, mReverseSolverTest.AcceptedRevision, budget)
            TestWriteReverseBinaryVector handle, mReverseSolverTest.ReadPolishEffortTest
            If solved Then
                mReverseSolverTest.CopyTrial field, actual
                If Not mReverseSolverTest.TryReadEnvelopeCoordinate(False, energy) Then Err.Raise 5, , "Нет H найденного корня."
                TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, actual
                Set record = New Collection: record.Add caseIndex: record.Add role: record.Add field: record.Add returned: records.Add record
            Else
                TestWriteReverseBinaryVector handle, Array(0#): TestWriteReverseBinaryVector handle, Array(0#)
            End If
            TestWriteReverseBinaryVector handle, Array(energy)
            Print #logHandle, "CASE;" & CStr(caseIndex) & ";ROLE;" & CStr(role) & ";SOLVED;" & CStr(solved) & ";LAMBDA;" & TestOnsetNumber(returned) & ";PRIMARY;" & CStr(used)
            Print #logHandle, "DETAILS;" & mReverseSolverTest.ResultMeta.DiagnosticDetails
            mReverseSolverTest.DiscardStep
            If mReverseSolverTest.AcceptedRevision <> revision Or used > 30 Then Err.Raise 5, , "H-пара изменила историю или превысила бюджет30."
        Next role
    Next caseIndex
    Close #handle: Close #logHandle
    Dim index As Long
    For index = 1 To records.Count
        Set record = records(index): field = record(3)
        TestExportSignedArcTrial "HCorner.Case" & CStr(record(1)) & "Role" & CStr(record(2)), field, CDbl(record(4))
    Next index
    TestRunHCornerPairs = "TOTAL_H_CORNER_PAIRS_TEST: passed=1; failed=0; actual numerical outcomes recorded; no history acceptance or physical capacity"
End Function''')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig');anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
left,tail=code.split(anchor);assert tail.startswith(prefix+'\n\n');tail=tail[len(prefix):]
code=left+anchor+prefix+'\n\n'.join(prefix+'\n\n'+e for e in entries)+'\n'+tail
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');compiler=replace(compiler,"-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledNewtonSolver','CCoupledArcContinuation','modTestCoupledElasticState') -EntryPointNames @('TestReadSharedCurveCases','TestRunHCornerPairs')")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';script=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig');anchor='        else { $result = [string]$excel.Run($macroName) }'
script=replace(script,anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected shared-curve setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadSharedCurveCases"))
        Save-Progress
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestRunHCornerPairs"))
        Save-Progress
''')
values[name]=script.replace('\n','\r\n').encode('utf-8-sig')
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector'];counter=json.loads((here/'ArcMetricCounterexample.json').read_bytes());cases=[]
for item in counter['cases']:
    d=item['initial_distance_to_peak'];h=item['step'];c=item['curvature'];peak=2+c*d*d
    ref=peak-c*d*d;full=peak-c*(h-d)**2;half=peak-c*(h/2-d)**2
    cases.append([h,0.,full,ref,abs(2*half-ref-full),item['alpha'],.005,1e-15])
cases+=[[1.,0.,2.,2.,0.,.0001,.005,1e-15],[1.,.006,2.,2.,0.,.0001,.005,1e-15]]
values['Input/SharedCurveCases.bin']=vector([len(cases)])+b''.join(vector(x) for x in cases)
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),
 'SourceArcMetricCounterexample.json':(here/'ArcMetricCounterexample.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():
    assert max(map(len,data.decode('utf-8-sig',errors='replace').splitlines()),default=0)<1024 or name.endswith('.bin'),name
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='SharedCurveHControlFromVerified20',actual_import_this_pass=['CCoupledNewtonSolver','CCoupledArcContinuation','modTestCoupledElasticState'],actual_component_count=256,
    source20_exact=True,shared_curve_metric_for_arc_return_and_H=True,two_relative_lambda_half_gates_replaced=True,
    principal_repeat_lambda_gate_unchanged=True,materials_and_force_tolerances_unchanged=True,
    three_independent_H_pairs_half_and_full_share30_primary_iterations=True,no_history_acceptance_or_physical_limit=True,
    analytical12_known_peak_cases_plus_plateau_and_field_error=True,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'COMMON CURVE METRIC;ARC/RETURN/H;14 ANALYTICAL CASES;THREE STRICT H PAIRS FROM TRUE20')
