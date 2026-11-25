"""ДЛЯ ТЕСТОВ. Сохраняет проверенное дуговое продолжение только после принятия точного кандидата и проверяет защиту его контекста."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold824AutomaticLoadingCold_20261125';root=base/'Cold827RetainedArcControl_20261125'
proof=json.loads((parent/'OperationalColdProof.json').read_bytes())
assert proof['accepted_states']==51 and proof['cold']['native']['own_process_closed']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'RetainedArcCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
anchor='Private mPostArcRefinedProbes As Long\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''Private mSignedArcField() As Double, mSignedArcProjection() As Double
Private mSignedArcExpectedRevision As Long, mSignedArcRetained As Long, mSignedArcNextRatio As Double
''',1)
anchor='    mPostArcRefinedProbes = 0\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''    Erase mSignedArcField: Erase mSignedArcProjection
    mSignedArcExpectedRevision = -1: mSignedArcRetained = 0: mSignedArcNextRatio = 1#
''',1)
anchor='        If increment < minimum Then\n            If Not unresolvedSteel Then arcRecovered = TryArcContinuationProbe(history, current, trial, nominal, increment, minimum)';assert code.count(anchor)==1
code=code.replace(anchor,'''        Dim retainArc As Boolean: retainArc = False
        If Not unresolvedSteel Then retainArc = CanRetainSignedArcControl()
        If retainArc Then
            arcRecovered = TryArcContinuationProbe(history, current, trial, nominal, increment, minimum)
            If arcRecovered Then mSignedArcRetained = mSignedArcRetained + 1
        End If
        If Not arcRecovered And increment < minimum Then
            If Not unresolvedSteel And Not retainArc Then arcRecovered = TryArcContinuationProbe(history, current, trial, nominal, increment, minimum)''',1)
anchor='        mSignedArcRevision = mSolver.AcceptedRevision: mSignedArcAttempts = 0: mSignedArcRatio = 1#\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'        If CanRetainSignedArcControl() Then mSignedArcRatio = mSignedArcNextRatio\n',1)
anchor='    mSignedArcRecovered = mSignedArcRecovered + 1\n';assert code.count(anchor)==1
code=code.replace(anchor,'''    mSignedArcField = field: mSignedArcProjection = mProjection
    mSignedArcExpectedRevision = mSolver.AcceptedRevision + 1
    mSignedArcNextRatio = GeomMin(2#, 0.8 * Sqr(0.005 / GeomMax(arc.FieldError, 0.00000001))) ' Новая дуга заново проходит полную и половинную пробы; отношение ограничено контрактом TryStep.
'''+anchor,1)
anchor='    If mSignedArcProbes > 0 Then Details = "; вызовов проверенной знаковой дуги " & CStr(mSignedArcProbes) & "; принятых направлений " & CStr(mSignedArcRecovered) & "; "\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'    If mSignedArcRetained > 0 Then Details = Details & "; продолжений принятой знаковой дуги " & CStr(mSignedArcRetained)\n',1)
code+='''
' Разрешает следующую дуговую пару только после принятия точного предыдущего кандидата с той же проекцией.
' Отмена, повторное восстановление истории и смена меры прекращают режим; совпадение не определяется округлением или хешем.
Private Function CanRetainSignedArcControl() As Boolean
    If mSignedArcRecovered = 0 Or Not mUsesFieldControl Or mUsesEnvelopeControl Then Exit Function
    If mSolver Is Nothing Then Exit Function
    If Not mSolver.UsesSpatialHistory Or mSolver.HasPrincipalBoundaries Then Exit Function
    If mSolver.AcceptedRevision <> mSignedArcExpectedRevision Then Exit Function
    Dim accepted() As Double, i As Long: mSolver.CopyAccepted accepted
    If UBound(accepted) <> UBound(mSignedArcField) Or UBound(mProjection) <> UBound(mSignedArcProjection) Then Exit Function
    For i = 1 To UBound(accepted)
        If accepted(i) <> mSignedArcField(i) Then Exit Function
    Next i
    For i = 1 To UBound(mProjection)
        If mProjection(i) <> mSignedArcProjection(i) Then Exit Function
    Next i
    CanRetainSignedArcControl = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет только точное совпадение контекста принятой дуги и независимость от чужой ревизии, поля или меры; равновесие не решает.
Friend Function TestRetainedArcContext(ByVal solver As CCoupledNewtonSolver, ByRef projection() As Double, ByVal mode As Long) As Variant
    Set mSolver = solver: mSolver.CopyAccepted mSignedArcField
    mProjection = projection: mSignedArcProjection = projection
    mSignedArcExpectedRevision = solver.AcceptedRevision: mSignedArcRecovered = 1
    mUsesFieldControl = True: mUsesEnvelopeControl = False
    If mode = 1 Then mSignedArcExpectedRevision = mSignedArcExpectedRevision - 1
    If mode = 2 Then mSignedArcField(1) = mSignedArcField(1) + 0.00000001
    If mode = 3 Then mProjection(1) = mProjection(1) + 0.00000001
    If mode = 4 Then mUsesEnvelopeControl = True
    If mode = 5 Then mUsesFieldControl = False
    If mode = 6 Then mSignedArcRecovered = 0
    Dim selected As Boolean: selected = CanRetainSignedArcControl()
    TestRetainedArcContext = Array(mode, Abs(CLng(selected)), solver.AcceptedRevision, mSignedArcExpectedRevision)
End Function
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='    TestRecentLoadingDirectionGuards\n    TestRunAutomaticLoadingWalk\n    TestRecentLoadingResetGuard\n';assert code.count(anchor)==1
code=code.replace(anchor,'    TestRunRetainedArcClassifier\n',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет совпадение принятой дуги и шесть отказов от чужого контекста без нового решения или принятия поля.
Private Sub TestRunRetainedArcClassifier()
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim before() As Double, after() As Double, projection() As Double, i As Long, mode As Long
    mReverseSolverTest.CopyAccepted before: projection = before
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\RetainedArc.Context.csv" For Output As #handle
    For mode = 0 To 6
        Dim subject As CCoupledDeformationContinuation: Set subject = New CCoupledDeformationContinuation
        Dim result As Variant: result = subject.TestRetainedArcContext(mReverseSolverTest, projection, mode)
        If CBool(result(1)) <> (mode = 0) Then Err.Raise 5, , "Неверно распознан точный контекст принятой дуги."
        If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Чтение контекста изменило принятую ревизию."
        mReverseSolverTest.CopyAccepted after
        For i = 1 To UBound(before)
            If before(i) <> after(i) Then Err.Raise 5, , "Чтение контекста изменило принятое поле."
        Next i
        Print #handle, "CONTEXT" & TestHardPreparedVector(result)
    Next mode
    Close #handle
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');imports=['CCoupledDeformationContinuation','modTestCoupledElasticState']
values.update({'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')',1).replace('\n','\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='RetainedVerifiedSignedArcControl',actual_import_this_pass=imports,actual_component_count=256,
    numerical_arc_Newton_materials_exact824=True,retention_requires_exact_accepted_field_projection_and_revision=True,
    next_ratio_bounded2_and_each_pair_verified_again=True,old_physical_steel_peak_and_global96_checks_preserved=True,
    context_classifier_without_new_NDS=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
