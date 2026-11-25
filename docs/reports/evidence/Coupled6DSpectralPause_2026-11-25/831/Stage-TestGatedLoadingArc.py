"""ДЛЯ ТЕСТОВ. Предпочитает дугу на новой границе нагружения и сохраняет её только после проверенного исходящего материального шага."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold827RetainedArcControl_20261125';root=base/'Cold831GatedLoadingArc_20261125'
assert json.loads((parent/'RetainedArcContextProof.json').read_bytes())['exact_context_positive_and_six_negative_controls']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'GatedLoadingCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
anchor='Private mSignedArcExpectedRevision As Long, mSignedArcRetained As Long, mSignedArcNextRatio As Double\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'Private mSignedArcOutgoing As Boolean\n',1)
anchor='    mSignedArcExpectedRevision = -1: mSignedArcRetained = 0: mSignedArcNextRatio = 1#\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'    mSignedArcOutgoing = False\n',1)
old='''        Dim retainArc As Boolean: retainArc = False
        If Not unresolvedSteel Then retainArc = CanRetainSignedArcControl()
        If retainArc Then
            arcRecovered = TryArcContinuationProbe(history, current, trial, nominal, increment, minimum)
            If arcRecovered Then mSignedArcRetained = mSignedArcRetained + 1
        End If
        If Not arcRecovered And increment < minimum Then
            If Not unresolvedSteel And Not retainArc Then arcRecovered = TryArcContinuationProbe(history, current, trial, nominal, increment, minimum)'''
new='''        Dim retainArc As Boolean, preferArc As Boolean: retainArc = False: preferArc = False
        If Not unresolvedSteel Then
            retainArc = CanRetainSignedArcControl()
            preferArc = retainArc
            If Not preferArc Then preferArc = CanStartRecentLoadingArc(history, current)
        End If
        If preferArc Then
            arcRecovered = TryArcContinuationProbe(history, current, trial, nominal, increment, minimum)
            If arcRecovered And retainArc Then mSignedArcRetained = mSignedArcRetained + 1
        End If
        If Not arcRecovered And increment < minimum Then
            If Not unresolvedSteel And Not preferArc Then arcRecovered = TryArcContinuationProbe(history, current, trial, nominal, increment, minimum)'''
assert code.count(old)==1;code=code.replace(old,new,1)
method,=re.findall(r'(?ms)^Private Function TryArcContinuationProbe\(.*?^End Function',code)
anchor='    If Not mSolver.UsesSpatialHistory Or mSolver.HasPrincipalBoundaries Or mDesignHistoryIndex = 0 Then Exit Function\n';assert method.count(anchor)==1
updated=method.replace(anchor,anchor+'    Dim continuingOutgoing As Boolean: continuingOutgoing = CanRetainSignedArcControl()\n',1)
anchor='    mSignedArcField = field: mSignedArcProjection = mProjection\n';assert updated.count(anchor)==1
updated=updated.replace(anchor,'''    Dim certificate As Variant, status As Variant
    certificate = arc.Verification: status = certificate(0)
    mSignedArcOutgoing = continuingOutgoing Or CLng(status(28)) = 4
'''+anchor,1)
code=code.replace(method,updated,1)
anchor='    If mSignedArcRecovered = 0 Or Not mUsesFieldControl Or mUsesEnvelopeControl Then Exit Function';assert code.count(anchor)==1
code=code.replace(anchor,'    If mSignedArcRecovered = 0 Or Not mSignedArcOutgoing Or Not mUsesFieldControl Or mUsesEnvelopeControl Then Exit Function',1)
anchor='    mSignedArcExpectedRevision = solver.AcceptedRevision: mSignedArcRecovered = 1\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'    mSignedArcOutgoing = True\n',1)
anchor='    If mode = 6 Then mSignedArcRecovered = 0\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'    If mode = 7 Then mSignedArcOutgoing = False\n',1)
old='    If mSignedArcRetained > 0 Then Details = Details & "; продолжений принятой знаковой дуги " & CStr(mSignedArcRetained)'
new='    If mSignedArcRetained > 0 Then Details = Details & "; продолжений после выхода на нагружение " & CStr(mSignedArcRetained) & "; "'
assert code.count(old)==1;code=code.replace(old,new,1)
code+='''
' Выбирает существующую материальную координату сразу на единственной недавно достигнутой границе, до дробления фиксированной меры.
' Номер станции не задаётся контроллером; ядро проверяет актуальную ревизию, поверхность и положительный прогноз.
Private Function CanStartRecentLoadingArc(ByVal history As Collection, ByRef current As TBranchPoint) As Boolean
    If mSolver Is Nothing Or history Is Nothing Or mTrace Is Nothing Then Exit Function
    If Not mSolver.UsesSpatialHistory Or mSolver.HasPrincipalBoundaries Or mDesignHistoryIndex = 0 Or mUsesEnvelopeControl Then Exit Function
    If history.Count < 4 Or current.HistoryIndex <> history.Count Then Exit Function
    Dim accepted() As Double, actual() As Double, previous() As Double, prediction() As Double, i As Long
    accepted = history(current.HistoryIndex): mSolver.CopyAccepted actual
    If UBound(accepted) <> UBound(actual) Then Exit Function
    For i = 1 To UBound(actual)
        If accepted(i) <> actual(i) Then Exit Function
    Next i
    previous = history(current.HistoryIndex - 1): prediction = accepted
    For i = 1 To mSolver.PhysicalDegreeCount: prediction(i) = 2# * accepted(i) - previous(i): Next i
    Dim info As Variant
    CanStartRecentLoadingArc = mSolver.TryReadRecentLoadingDeparture(prediction, info)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет ранний выбор существующей границы и отказ от непригодного контекста без решения или изменения принятого материала.
Friend Function TestProactiveLoadingContext(ByVal solver As CCoupledNewtonSolver, ByVal fields As Collection, _
        ByVal trace As CCoupledLoadTrace, ByVal mode As Long) As Variant
    Set mSolver = solver: Set mTrace = trace
    mDesignHistoryIndex = 1: mUsesEnvelopeControl = False
    Dim current As TBranchPoint, history As New Collection, field() As Double, i As Long
    For i = 1 To fields.Count
        field = fields(i)
        If mode = 3 And i = fields.Count Then field(1) = field(1) + 0.00000001
        If mode = 6 And i = fields.Count - 1 Then field = fields(fields.Count)
        history.Add field
    Next i
    current.HistoryIndex = history.Count
    If mode = 1 Then mDesignHistoryIndex = 0
    If mode = 2 Then mUsesEnvelopeControl = True
    If mode = 4 Then current.HistoryIndex = current.HistoryIndex - 1
    If mode = 5 Then Set mTrace = Nothing
    Dim selected As Boolean: selected = CanStartRecentLoadingArc(history, current)
    TestProactiveLoadingContext = Array(mode, Abs(CLng(selected)), solver.AcceptedRevision)
End Function
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Private Sub TestRunRetainedArcClassifier\(\).*?^End Sub',code)
assert method.count('    For mode = 0 To 6\n')==1
code=code.replace(method,method.replace('    For mode = 0 To 6\n','    For mode = 0 To 7\n',1),1)
anchor='    TestRunRetainedArcClassifier\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'    TestRunProactiveLoadingClassifier\n',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет выбор новой принятой границы и шесть отказов от контекста при точном исходном43-точечном пути.
Private Sub TestRunProactiveLoadingClassifier()
    Dim before() As Double, after() As Double, revision As Long, mode As Long, i As Long
    mReverseSolverTest.CopyAccepted before: revision = mReverseSolverTest.AcceptedRevision
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\ProactiveLoading.Context.csv" For Output As #handle
    For mode = 0 To 6
        Dim subject As CCoupledDeformationContinuation: Set subject = New CCoupledDeformationContinuation
        Dim result As Variant
        result = subject.TestProactiveLoadingContext(mReverseSolverTest, mReverseFieldsTest, mReverseTraceTest, mode)
        If CBool(result(1)) <> (mode = 0) Then Err.Raise 5, , "Неверный выбор раннего выхода на материальную координату."
        If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Ранний выбор изменил принятую ревизию."
        mReverseSolverTest.CopyAccepted after
        For i = 1 To UBound(before)
            If before(i) <> after(i) Then Err.Raise 5, , "Ранний выбор изменил принятое поле."
        Next i
        Print #handle, "PROACTIVE" & TestHardPreparedVector(result)
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
prep['directed_scope']=dict(case='GatedOutgoingArcAndProactiveUniqueBoundary',actual_import_this_pass=imports,actual_component_count=256,
    numerical_arc_Newton_materials_exact827=True,retention_starts_only_after_verified_loading_departure_kind4=True,
    proactive_coordinate_only_at_solver_selected_unique_recent_boundary=True,
    initial_preevent_approach_does_not_start_retention=True,all_physical_and_global30_96_checks_preserved=True,
    context_classifiers_without_new_NDS=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
