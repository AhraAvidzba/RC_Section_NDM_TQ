"""ДЛЯ ТЕСТОВ. Встраивает постоянную дуговую метрику и исходящую касательную в общие классы частной проверяемой книги."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent
base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold857SpatialDeclineRecovery_20261125'
root=base/'Cold867FixedMetricKernel_20261126'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'FixedMetricCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}

def replace(code,old,new):
    """Останавливает подготовку при несовпадении единственного ожидаемого участка."""
    assert code.count(old)==1,(old[:100],code.count(old))
    return code.replace(old,new,1)

name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
code=replace(code,'Private mHybridActive As Boolean, mHybridScalesReady As Boolean',
'''Private mArcScaleReady As Boolean, mArcScale As Double, mArcScaleDirection() As Double
Private mHybridActive As Boolean, mHybridScalesReady As Boolean''')
code=replace(code,'    mHybridScalesReady = False: Erase mHybridScales',
'    mHybridScalesReady = False: Erase mHybridScales\n    mArcScaleReady = False: Erase mArcScaleDirection')
start=code.index('Friend Function BeginNewtonLineTest(');end=code.index('End Function',start)+len('End Function')
old=code[start:end]
body=old.replace('Friend Function BeginNewtonLineTest(', 'Private Function BeginControlLinearization(').replace('Не завершена предыдущая тестовая проба.','Не завершена предыдущая проба равновесия.').replace('    BeginNewtonLineTest = True','    BeginControlLinearization = True')
code=code[:start]+'''Friend Function BeginNewtonLineTest(ByRef baseLoad() As Double, ByRef direction() As Double, _
        ByRef projection() As Double, ByVal target As Double, ByRef seed() As Double, ByVal factor As Double) As Boolean
    BeginNewtonLineTest = BeginControlLinearization(baseLoad, direction, projection, target, seed, factor)
End Function'''+code[end:]
code+='''

' Открывает неизменяемую систему для чтения касательной; кандидат и принятая история не публикуются.
'''+body+'''

' Закрывает только читающий контекст; ни материал, ни принятую ревизию не изменяет.
Private Sub EndControlLinearization()
    mHybridActive = False: mControlActive = False
End Sub

' Возвращает постоянный упругий масштаб нагрузки в тензорной метрике подготовленного сечения.
' Кэш хранит только число и точное направление; при Initialize геометрия и материалы инвалидируют его.
Friend Function TryElasticArcScale(ByRef direction() As Double, ByRef alpha As Double) As Boolean
    Dim i As Long, exact As Boolean: exact = mArcScaleReady
    If exact Then
        For i = 1 To COMPONENT_COUNT
            If direction(i) <> mArcScaleDirection(i) Then exact = False: Exit For
        Next i
    End If
    If exact Then alpha = mArcScale: TryElasticArcScale = True: Exit Function
    Dim deformation() As Double, recovered() As Double, stress() As Double, raw As Variant
    If Not mSection.Solve(direction(1), direction(2), direction(3), direction(4), direction(5), direction(6), _
            deformation, recovered, stress, raw) Then Exit Function
    Dim field() As Double, zero() As Double
    ReDim field(1 To mDegreeCount): ReDim zero(1 To mDegreeCount)
    For i = 1 To mDegreeCount: field(i) = raw(i): Next i
    alpha = SpatialFieldDistance(zero, field)
    If alpha <= 0# Then Exit Function
    mArcScale = alpha: mArcScaleDirection = direction: mArcScaleReady = True
    TryElasticArcScale = True
End Function

' Строит исходящую производную поля по множителю на малом положительном продолжении принятой секущей.
' Пробное поле выбирает активную ветвь материала; оно не является равновесием и не принимается.
Friend Function TrySpatialLoadTangent(ByRef baseLoad() As Double, ByRef direction() As Double, _
        ByRef previous() As Double, ByRef accepted() As Double, ByVal factor As Double, _
        ByRef tangent() As Double, ByRef probe() As Double) As Boolean
    If Not mUseSpatialHistory Or mCandidateReady Or mHybridActive Then Exit Function
    Dim revision As Long: revision = mAcceptedRevision
    Dim i As Long, j As Long, projection() As Double, zero() As Double, residual() As Double, matrix() As Double
    probe = accepted: ReDim projection(1 To mDegreeCount): ReDim zero(1 To mDegreeCount + 1)
    For i = 1 To mDegreeCount: probe(i) = accepted(i) + (accepted(i) - previous(i)) / 1024#: Next i
    On Error GoTo Failed
    If Not BeginControlLinearization(baseLoad, direction, projection, 0#, probe, factor) Then Exit Function
    If Not TryReadBranchLinearization(zero, residual, matrix) Then GoTo Finished
    Dim system As New CCoupledLinearSystem, physical() As Double, rhs() As Double
    ReDim physical(1 To mDegreeCount, 1 To mDegreeCount): ReDim rhs(1 To mDegreeCount)
    For i = 1 To mDegreeCount
        rhs(i) = -matrix(mDegreeCount + 1, i) / mHybridScales(mDegreeCount + 1)
        For j = 1 To mDegreeCount: physical(j, i) = matrix(j, i): Next j
    Next i
    If Not system.Solve(physical, rhs, mDegreeCount) Then GoTo Finished
    ReDim tangent(1 To mDegreeCount)
    For i = 1 To mDegreeCount: tangent(i) = rhs(i) * mHybridScales(i): Next i
    TrySpatialLoadTangent = True
Finished:
    EndControlLinearization
    If mAcceptedRevision <> revision Then Err.Raise 5, "CCoupledNewtonSolver", "Чтение касательной изменило принятую историю."
    Exit Function
Failed:
    Dim number As Long, description As String: number = Err.Number: description = Err.Description
    EndControlLinearization
    Err.Raise number, "CCoupledNewtonSolver.TrySpatialLoadTangent", description
End Function
'''
code=replace(code,'Friend Sub EndNewtonLineTest()\n    mHybridActive = False: mControlActive = False\nEnd Sub',
             'Friend Sub EndNewtonLineTest()\n    EndControlLinearization\nEnd Sub')
values[name]=code

name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
code=replace(code,'Private mNextStepRatioLimit As Double','''Private mNextStepRatioLimit As Double
Private mReferenceScale As Double, mJointError As Double, mStepError As Double
Private mUsesTangent As Boolean, mPredictionVerification As Variant''')
code=replace(code,'    mMaximumIterations = maximumIterations: mModulus = solver.ReferenceModulus: mNextStepRatioLimit = 2#',
'''    mMaximumIterations = maximumIterations: mModulus = solver.ReferenceModulus: mNextStepRatioLimit = 2#
    If Not solver.TryElasticArcScale(direction, mReferenceScale) Then _
        Err.Raise 5, "CCoupledArcContinuation", "Не определён постоянный упругий масштаб выбранного LoadPath."''')
code=replace(code,'    mCalls = 0: mPrimary = 0: mError = -1#: mDetails = vbNullString: mCertificate = Empty',
'''    mCalls = 0: mPrimary = 0: mError = -1#: mDetails = vbNullString: mCertificate = Empty
    mJointError = -1#: mStepError = -1#: mUsesTangent = False: mPredictionVerification = Empty''')
start=code.index('    For i = 1 To 3\n        If factor(i + 1) = factor(i) Then Exit Function');end=code.index('    If Not loadingDeparture Then\n',start)
code=code[:start]+'''    Dim alpha As Double, dl As Double, secant As Double, radius As Double, q As Double, target As Double
    alpha = mReferenceScale: dl = factor(4) - factor(3)
    Dim loadingPrediction() As Double: loadingPrediction = accepted
    For i = 1 To mSolver.PhysicalDegreeCount: loadingPrediction(i) = accepted(i) + ratio * movement(i): Next i
    Dim recent As Variant, loadingDeparture As Boolean
    loadingDeparture = mSolver.TryReadRecentLoadingDeparture(loadingPrediction, recent)
    If Not loadingDeparture And callAllowance >= 4 Then
        Dim tangent() As Double, probe() As Double, zero() As Double, tangentLength As Double, orientation As Double
        mCalls = mCalls + 1
        If mSolver.TrySpatialLoadTangent(mBase, mDirection, previous, accepted, factor(4), tangent, probe) Then
            ReDim zero(1 To UBound(accepted))
            tangentLength = mSolver.SpatialFieldDistance(zero, tangent)
            orientation = alpha * alpha * dl
            For i = 1 To UBound(projection): orientation = orientation + projection(i) * tangent(i) * spacing: Next i
            If tangentLength > 0# And orientation <> 0# Then
                Dim tangentScale As Double: tangentScale = Sgn(orientation) * spacing / tangentLength
                For i = 1 To UBound(movement): movement(i) = tangentScale * tangent(i): Next i
                dl = tangentScale
                Dim destination() As Double: destination = accepted
                For i = 1 To UBound(destination): destination(i) = accepted(i) + movement(i): Next i
                If Not mSolver.SelectSpatialDirectionControl(accepted, destination, movement, projection, coordinate, spacing) Then Exit Function
                mUsesTangent = True
                mPredictionVerification = Array(Array(1# / 1024#, dl, tangentLength, orientation, alpha), probe, tangent, movement)
            End If
        End If
    End If
    secant = Sqr(spacing * spacing + alpha * alpha * dl * dl)
    If secant <= 0# Then Exit Function
    radius = ratio * secant: q = alpha * alpha * dl / secant: target = q * factor(4)
    For i = 1 To UBound(projection)
        projection(i) = projection(i) * spacing / secant: target = target + projection(i) * accepted(i)
    Next i
    Dim revision As Long: revision = mSolver.AcceptedRevision
    Dim halfSeed() As Double, fullSeed() As Double, half() As Double, full() As Double, ah() As Double, af() As Double
    halfSeed = accepted: fullSeed = accepted
    For i = 1 To mSolver.PhysicalDegreeCount
        halfSeed(i) = accepted(i) + 0.5 * ratio * movement(i): fullSeed(i) = accepted(i) + ratio * movement(i)
    Next i
    Dim sh As Boolean, sf As Boolean, lh As Double, lf As Double, halfSteps As Long, pairSteps As Long, prediction As Variant
    Dim fullPrediction As Double: fullPrediction = factor(4) + ratio * dl
'''+code[end:]
code=replace(code,'            lambdaBound = GeomMax(FIELD_ERROR_LIMIT * Abs(lf - factor(4)), 8# * mRelative * (1# + Abs(factor(4))))',
'            lambdaBound = PairLambdaBound(length, mError, lf, factor(4), lambdaError)')
code=replace(code,'    lambdaBound = GeomMax(FIELD_ERROR_LIMIT * Abs(factor - reference), 8# * mRelative * (1# + Abs(reference)))',
'    lambdaBound = PairLambdaBound(spacing, error, factor, reference, Abs(2# * halfLambda - reference - factor))') if code.count('    lambdaBound = GeomMax(FIELD_ERROR_LIMIT * Abs(factor - reference), 8# * mRelative * (1# + Abs(reference)))')==1 else code
# Три оставшиеся проверки имеют разные имена длины и ошибки; замена привязана к месту, а не к номеру строки.
old='lambdaBound = GeomMax(FIELD_ERROR_LIMIT * Abs(factor - reference), 8# * mRelative * (1# + Abs(reference)))'
for marker,new in [('    error = mSolver.SpatialFieldDistance(face, doubled) / spacing','PairLambdaBound(spacing, error, factor, reference, Abs(2# * halfLambda - reference - factor))'),
 ('            lambdaError = Abs(2# * probeLambda - reference - factor)','PairLambdaBound(length, error, factor, reference, lambdaError)'),
 ('    Dim lambdaBound As Double, advance As Double, oldLength As Double','PairLambdaBound(length, mError, factor, reference, Abs(2# * halfLambda - reference - factor))')]:
    pos=code.index(marker);pos2=code.index(old,pos);code=code[:pos2]+code[pos2:].replace(old,'lambdaBound = '+new,1)
assert old not in code
code=replace(code,'mError, lambdaError, lambdaBound, candidateLambda, eventHalfLambda, eventError, Abs(CLng(eventSolved)), eventLambda, selectedKind, mApproaches.Count)',
'mError, lambdaError, lambdaBound, candidateLambda, eventHalfLambda, eventError, Abs(CLng(eventSolved)), eventLambda, selectedKind, mApproaches.Count, _\n        mJointError, alpha, Abs(CLng(mUsesTangent)), mStepError)')
code=replace(code,'    If unloading Then mDetails = mDetails &',
'    mDetails = mDetails & "; совместная ошибка " & CStr(mJointError) & "; исходящая касательная " & CStr(mUsesTangent)\n    If unloading Then mDetails = mDetails &')
code+='''

' Ограничивает кривизну пары в постоянной метрике поля и нагрузки, сохраняя отдельный предел ошибки поля.
' Абсолютный шум множителя учитывает прежнюю точность сил; изменение lambda не используется как делитель у экстремума.
Private Function PairLambdaBound(ByVal length As Double, ByVal fieldError As Double, ByVal factor As Double, _
        ByVal reference As Double, ByVal lambdaError As Double) As Double
    Dim radius As Double, fieldDefect As Double, allowance As Double, noise As Double
    radius = Sqr(length * length + mReferenceScale * mReferenceScale * (factor - reference) ^ 2)
    If radius <= 0# Then Exit Function
    fieldDefect = fieldError * length
    allowance = (FIELD_ERROR_LIMIT * radius) ^ 2 - fieldDefect * fieldDefect
    noise = 8# * mRelative * (1# + Abs(reference))
    PairLambdaBound = noise
    If allowance > 0# Then PairLambdaBound = GeomMax(noise, Sqr(allowance) / mReferenceScale)
    mJointError = Sqr(fieldDefect * fieldDefect + (mReferenceScale * lambdaError) ^ 2) / radius
    Dim resolvedLambdaDefect As Double: resolvedLambdaDefect = lambdaError
    If lambdaError <= noise Then resolvedLambdaDefect = 0#
    mStepError = GeomMax(fieldError, Sqr(fieldDefect * fieldDefect + (mReferenceScale * resolvedLambdaDefect) ^ 2) / radius)
End Function

' Возвращает обе составляющие кривизны для адаптации шага; фактическая ошибка поля доступна отдельно.
Friend Property Get StepError() As Double
    StepError = mStepError
End Property

' Указывает, что прогноз выбран по исходящей касательной без навязывания роста нагрузки.
Friend Property Get UsesTangentPredictor() As Boolean
    UsesTangentPredictor = mUsesTangent
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает независимую копию вычисленной производной и пробного поля; данные не возвращаются в алгоритм.
Friend Function TestPredictionVerification() As Variant
    TestPredictionVerification = mPredictionVerification
End Function
'''
values[name]=code

name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
code=replace(code,'    If mSignedArcAttempts >= 4 Then Exit Function','    If mSignedArcAttempts >= 6 Then Exit Function \' Шесть долей включают 1/16, необходимую для проверки кривизны у экстремума; общий бюджет не меняется.')
code=replace(code,'    mSignedArcOutgoing = continuingOutgoing Or CLng(status(28)) = 4','    mSignedArcOutgoing = continuingOutgoing Or arc.UsesTangentPredictor Or CLng(status(28)) = 4')
code=replace(code,'GeomMax(arc.FieldError, 0.00000001)','GeomMax(arc.StepError, 0.00000001)')
values[name]=code

name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
start=code.index('Private Sub TestRunAutomaticLoadingWalk()');end=code.index('End Sub',start)+len('End Sub')
entry=code[start:end].replace('Private Sub TestRunAutomaticLoadingWalk()', 'Public Sub TestRunFixedMetricArcWalk()').replace('attemptsAtRevision >= 4','attemptsAtRevision >= 6').replace('arc.FieldError, 0.00000001','arc.StepError, 0.00000001').replace('TestWriteAutomaticHGuideWalk attempt, arc','TestWriteFixedMetricGuide attempt, arc')
entry+='''

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает версию3 диагностического ориентира: H, пластическое нагружение и исходящую касательную.
Public Sub TestWriteFixedMetricGuide(ByVal attempt As Long, ByVal arc As CCoupledArcContinuation)
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\LoadingWalk.Guide" & Format$(attempt, "000") & ".bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMWALKH00000003"
    Put #handle, , magic
    Dim guide As Variant, vector As Variant
    TestWriteReverseBinaryVector handle, Array(attempt, arc.NextStepRatioLimit, arc.StepError)
    guide = arc.EnvelopeGuideVerification
    TestWriteReverseBinaryVector handle, Array(Abs(CLng(Not IsEmpty(guide))))
    If Not IsEmpty(guide) Then
        For Each vector In guide: TestWriteReverseBinaryVector handle, vector: Next vector
    End If
    guide = arc.LoadingGuideVerification
    TestWriteReverseBinaryVector handle, Array(Abs(CLng(Not IsEmpty(guide))))
    If Not IsEmpty(guide) Then
        For Each vector In guide: TestWriteReverseBinaryVector handle, vector: Next vector
    End If
    guide = arc.TestPredictionVerification
    TestWriteReverseBinaryVector handle, Array(Abs(CLng(Not IsEmpty(guide))))
    If Not IsEmpty(guide) Then
        For Each vector In guide: TestWriteReverseBinaryVector handle, vector: Next vector
    End If
    Close #handle
End Sub
'''
# Воспроизводим позицию AddFromString: после деклараций, до первого метода.
anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n'
entries=[]
for method in ('TestRunFixedMetricArcWalk','TestWriteFixedMetricGuide'):
    match=re.search(r'(?ms)^Public Sub '+method+r'\(.*?^End Sub\b',entry);assert match
    entries.append("' FOR TESTS: baseline entry point only; existing methods retained.\n"+match[0])
code=replace(code,anchor,anchor+'\n\n'+'\n\n'.join(entries)+'\n')
values[name]=code
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
compiler=replace(compiler,"-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledNewtonSolver','CCoupledArcContinuation','CCoupledDeformationContinuation','modTestCoupledElasticState') -EntryPointNames @('TestRunFixedMetricArcWalk','TestWriteFixedMetricGuide')")
values['Compile-TestExternalPoint.ps1']=compiler
name='tools/build_workbook/Update-IsolatedOptimizerCode.ps1';code=(root/name).read_text('utf-8-sig')
code=replace(code,"if ($EntryPointNames.Count -and ($ModuleNames.Count -ne 1 -or $RetiredModules.Count -or $TraceGeometry)) {throw 'Entry-point instrumentation requires exactly one module and no retirement/trace.'}",
"if ($EntryPointNames.Count -and ('modTestCoupledElasticState' -notin $ModuleNames -or $RetiredModules.Count -or $TraceGeometry)) {throw 'Entry-point instrumentation requires the test module and no retirement/trace.'}")
assert code.count('if ($EntryPointNames.Count) {')==1 and code.count('if($EntryPointNames.Count){')==1
code=code.replace('if ($EntryPointNames.Count) {',"if ($EntryPointNames.Count -and $file.BaseName -eq 'modTestCoupledElasticState') {")
code=code.replace('if($EntryPointNames.Count){',"if($EntryPointNames.Count -and $file.BaseName -eq 'modTestCoupledElasticState'){")
values[name]=code
name='tools/build_workbook/Run-Audit02ModeTests.ps1';code=(root/name).read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }'
code=replace(code,anchor,anchor+'''
        # ДЛЯ ТЕСТОВ: общий рабочий алгоритм восстанавливает и продолжает настоящую историю19.
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected fixed-metric setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestRunFixedMetricArcWalk"))
        Save-Progress
''')
values[name]=code
for name,text in values.items():
    assert max(map(len,text.splitlines()))<1024,name
    data=text.replace('\n','\r\n').encode('utf-8-sig');(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
for name,data in {
    'Input/Reverse1.bin':(base/'Cold866Cold19ReferenceMetricPairs_20261126/Input/Reverse1.bin').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),
    Path(__file__).name:Path(__file__).read_bytes()}.items():
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='GeneralFixedMetricKernelFromCold19',actual_import_this_pass=['CCoupledNewtonSolver','CCoupledArcContinuation','CCoupledDeformationContinuation','modTestCoupledElasticState'],actual_component_count=256,
    parent=str(parent),working_classes_changed=True,original_force_tolerance=True,unchanged_materials_and_constitutive_history=True,
    constant_elastic_reference_metric=True,read_only_refined_outgoing_tangent=True,attempts_per_revision=6,maximum_calls=96,
    physical_capacity_certified=False,whole_goal_completed=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'GENERAL WORKING KERNEL CHANGED;FOUR IMPORTS;NATIVE VERIFICATION PENDING')
