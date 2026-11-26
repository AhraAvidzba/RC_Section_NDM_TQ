"""ДЛЯ ТЕСТОВ. Даёт активной пространственной ветви одну свежую дуговую попытку на принятую ревизию, сохраняя точный контракт удержания кэша."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold888StrictEnvelopeGrowth_20261126';source=base/'Cold893Arc56Diagnostics_20261126';root=base/'Cold895FreshSpatialArc_20261126'
assert all(c['solved'] for c in json.loads((source/'Arc56DiagnosticsProof.json').read_bytes())['cases'])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'FreshArcCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
def replace(code,old,new):
    """Проверяет однозначный контракт перед изменением общего контроллера."""
    assert code.count(old)==1,(old[:90],code.count(old));return code.replace(old,new,1)
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
anchor='Private mSignedArcExpectedRevision As Long, mSignedArcRetained As Long, mSignedArcNextRatio As Double'
code=replace(code,anchor,anchor+'\nPrivate mSignedArcFreshStarts As Long')
anchor='    mSignedArcExpectedRevision = -1: mSignedArcRetained = 0: mSignedArcNextRatio = 1#'
code=replace(code,anchor,anchor+'\n    mSignedArcFreshStarts = 0')
anchor='            If Not preferArc Then preferArc = CanStartRecentLoadingArc(history, current)'
code=replace(code,anchor,anchor+'''
            If Not preferArc Then
                preferArc = CanStartFreshSpatialArc(history, current)
                If preferArc Then mSignedArcFreshStarts = mSignedArcFreshStarts + 1
            End If''')
anchor="' Разрешает следующую дуговую пару только после принятия точного предыдущего кандидата с той же проекцией."
method=''' ' Разрешает одну заново построенную дуговую пробу на принятой армированной пространственной ветви.
' Старые поле, проекция и касательная не используются: TryStep заново готовит их по текущей истории и проверяет полную/половинную пару.
Private Function CanStartFreshSpatialArc(ByVal history As Collection, ByRef current As TBranchPoint) As Boolean
    If mSignedArcRecovered = 0 Or Not mUsesFieldControl Or Not mBranchTangentControl Then Exit Function
    If mUsesEnvelopeControl Or mUsesYieldControl Then Exit Function
    If mSolver Is Nothing Or mReinforcement Is Nothing Then Exit Function
    If Not mReinforcement.HasSteel Or Not mSolver.UsesSpatialHistory Or mSolver.HasPrincipalBoundaries Then Exit Function
    If mDesignHistoryIndex = 0 Or history Is Nothing Or mTrace Is Nothing Then Exit Function
    If history.Count < 4 Or current.HistoryIndex <> history.Count Or mTrace.Count <> history.Count Then Exit Function
    If mSignedArcRevision = mSolver.AcceptedRevision And mSignedArcAttempts > 0 Then Exit Function
    Dim fraction As Double, factor As Double
    mTrace.ReadPoint history.Count, fraction, factor
    If fraction <> 1# Or factor <> current.Lambda Or factor <= 0# Then Exit Function
    Dim accepted() As Double, stored() As Double, i As Long
    mSolver.CopyAccepted accepted: stored = history(history.Count)
    If UBound(accepted) <> UBound(stored) Then Exit Function
    For i = 1 To UBound(accepted)
        If accepted(i) <> stored(i) Then Exit Function
    Next i
    CanStartFreshSpatialArc = True
End Function

'''.lstrip()
code=replace(code,anchor,method+anchor)
anchor='    If mArcRecoveryTried Then Details = Details & "; вспомогательных проб псевдодуги и повтора: " & CStr(mArcRecoveryProbes)'
code=replace(code,anchor,anchor+'\n    If mSignedArcFreshStarts > 0 Then Details = Details & "; свежих дуговых проб на принятой пространственной ветви: " & CStr(mSignedArcFreshStarts)')
code+='''

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет начало свежей дуговой пробы при правильной истории, отказе на той же ревизии и чужих данных; материал не решает и не принимает.
Friend Function TestFreshSpatialArcSelection(ByVal solver As CCoupledNewtonSolver, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByVal fields As Collection, ByVal trace As CCoupledLoadTrace, ByVal mode As Long) As Variant
    Set mSolver = solver: Set mReinforcement = reinforcement
    Set mTrace = trace.CopyPrefix(trace.Count)
    mSignedArcRecovered = 1: mUsesFieldControl = True: mBranchTangentControl = True: mDesignHistoryIndex = 1
    mUsesEnvelopeControl = False: mUsesYieldControl = False: mSignedArcRevision = -1: mSignedArcAttempts = 0
    Dim current As TBranchPoint, fraction As Double, history As New Collection, i As Long, field() As Double
    For i = 1 To fields.Count: history.Add fields(i): Next i
    current.HistoryIndex = history.Count: mTrace.ReadPoint history.Count, fraction, current.Lambda
    Select Case mode
        Case 1: mSignedArcRecovered = 0
        Case 2: mBranchTangentControl = False
        Case 3: mUsesEnvelopeControl = True
        Case 4: mSignedArcRevision = solver.AcceptedRevision: mSignedArcAttempts = 1
        Case 5
            field = history(history.Count): field(1) = field(1) + 0.000000000001
            history.Remove history.Count: history.Add field
        Case 6
            Set mTrace = trace.CopyPrefix(trace.Count - 1): mTrace.Add 0.9, current.Lambda
        Case 7: Set mReinforcement = Nothing
        Case 8: mUsesYieldControl = True
        Case 9
            solver.CopyAccepted mSignedArcField: mSignedArcField(1) = mSignedArcField(1) + 1#
            ReDim mSignedArcProjection(1 To UBound(mSignedArcField)): mSignedArcProjection(1) = 99#
        Case 10: mSignedArcRevision = solver.AcceptedRevision - 1: mSignedArcAttempts = 1
    End Select
    Dim revision As Long: revision = solver.AcceptedRevision
    Dim selected As Boolean: selected = CanStartFreshSpatialArc(history, current)
    If solver.AcceptedRevision <> revision Then Err.Raise 5, , "Проверка нового контекста приняла материал."
    TestFreshSpatialArcSelection = Array(mode, Abs(CLng(selected)), revision, solver.AcceptedRevision)
End Function
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
entry='''Public Function TestReadFreshSpatialArcCases() As String
    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Проверяет допустимость нового контекста; чужие старые поля кэша не используются, а чужая текущая история отклоняется.
    Dim passed As Long, failed As Long, report As String, mode As Long, expected As Boolean, result As Variant
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\FreshSpatialArcCases.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMFRESHARC00001"
    Put #handle, , magic
    For mode = 0 To 10
        Dim controller As New CCoupledDeformationContinuation
        result = controller.TestFreshSpatialArcSelection(mReverseSolverTest, mReverseReinforcedTest, mReverseFieldsTest, mReverseTraceTest, mode)
        expected = (mode = 0 Or mode = 9 Or mode = 10)
        TestRecord CBool(result(1)) = expected, "fresh spatial arc eligibility mode " & CStr(mode), passed, failed, report
        TestRecord result(2) = result(3), "fresh spatial arc leaves material revision mode " & CStr(mode), passed, failed, report
        TestWriteReverseBinaryVector handle, result
        Set controller = Nothing
    Next mode
    Close #handle
    TestReadFreshSpatialArcCases = report & "TOTAL_FRESH_SPATIAL_ARC_CASES_TEST: passed=" & passed & "; failed=" & failed
End Function
'''
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig');anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+"' FOR TESTS: baseline entry point only; existing methods retained.\n\n\n"+entry,1);values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledDeformationContinuation','modTestCoupledElasticState') -EntryPointNames 'TestReadFreshSpatialArcCases'",1)
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig');anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected fresh-arc setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadFreshSpatialArcCases"))
        Save-Progress
''',1);values[name]=runner.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Input/Reverse1.bin':(source/'Input/Reverse1.bin').read_bytes(),'Input/SourceColdRawHistory.bin':(source/'Input/SourceColdRawHistory.bin').read_bytes(),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='FreshSpatialArcContext',actual_import_this_pass=['CCoupledDeformationContinuation','modTestCoupledElasticState'],actual_component_count=256,
    exact_old_retained_cache_gate_unchanged=True,one_fresh_try_per_accepted_revision=True,existing_arc_rebuilds_all_directions=True,
    only_after_verified_reinforced_spatial_arc=True,global96_and_primary30_unchanged=True,material_and_force_tolerances_unchanged=True,
    physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'ONE FRESH ARC TRY PER NEW ACCEPTED SPATIAL CONTEXT;OLD CACHE CONTRACT EXACT;SAME96/30')
