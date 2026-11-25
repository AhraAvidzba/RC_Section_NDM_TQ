"""ДЛЯ ТЕСТОВ. Сохраняет смысл исходной неудачной пробы после вспомогательной дуги и проверяет область действия кэша отказа."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import re

here=Path(__file__).resolve().parent; base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold785OperationalArcController_20261125'; root=base/'Cold786ArcFailureContext_20261125'
assert json.loads((parent/'OperationalArcProof.json').read_bytes())['operational_class_called_without_test_search_driver']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ArcFailureContextCompilationAncestor'; archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Private Function PrepareHybridRecovery\(.*?^End Function',code)
anchor='''    If lastProbeFailed Then
        code = mSolver.ResultMeta.ResultCode
        target = acceptedCoordinate + 2# * increment
'''
assert method.count(anchor)==1
updated=method.replace(anchor,'''    If lastProbeFailed Then
        code = mSolver.ResultMeta.ResultCode
        If mFailedProbeRevision = mSolver.AcceptedRevision Then
            sameProjection = True
            For i = 1 To UBound(mProjection)
                If mFailedProbeProjection(i) <> mProjection(i) Then sameProjection = False: Exit For
            Next i
            If sameProjection Then code = mFailedProbeCode ' Вспомогательные пробы не заменяют причину отказа той же исходной цели.
        End If
        target = acceptedCoordinate + 2# * increment
''',1)
code=code.replace(method,updated,1)
code=code.replace('CStr(mSignedArcRecovered)\n    Details = Details & "Мера деформации:', 'CStr(mSignedArcRecovered) & "; "\n    Details = Details & "Мера деформации:',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет выбор причины исходного отказа при совпавшем контексте и отказ от чужой проекции либо ревизии; НДС не решает.
Friend Function TestCachedArcFailureRecovery(ByVal solver As CCoupledNewtonSolver, ByRef projection() As Double, ByVal mode As Long) As Variant
    Set mSolver = solver: mProjection = projection: mFailedProbeProjection = projection
    mFailedProbeRevision = solver.AcceptedRevision: mFailedProbeCode = rcNumericalFailure
    mLastHybridRevision = -1: mHybridRecoveryProbes = 0
    If mode = 1 Then mFailedProbeProjection(1) = mFailedProbeProjection(1) + 1#
    If mode = 2 Then mFailedProbeRevision = mFailedProbeRevision + 1
    Dim increment As Double: increment = 0.1
    Dim selected As Boolean: selected = PrepareHybridRecovery(True, False, 1#, increment)
    TestCachedArcFailureRecovery = Array(Abs(CLng(selected)), mode, solver.ResultMeta.ResultCode, mFailedProbeCode, mHybridRecoveryProbes, increment)
End Function
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='    TestRunOperationalArc\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'    TestRunArcFailureContext\n',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт успешный текущий статус от принятого поля и проверяет маршрутизацию предшествовавшего отказа без изменения истории.
Private Sub TestRunArcFailureContext()
    Dim accepted() As Double, previous() As Double, movement() As Double, projection() As Double, coordinate As Double, spacing As Double, i As Long
    accepted = mReverseFieldsTest(mReverseFieldsTest.Count): previous = mReverseFieldsTest(mReverseFieldsTest.Count - 1): movement = accepted
    For i = 1 To UBound(movement): movement(i) = accepted(i) - previous(i): Next i
    If Not mReverseSolverTest.SelectSpatialDirectionControl(previous, accepted, movement, projection, coordinate, spacing) Then Err.Raise 5, , "Нет исходной меры контекстного теста."
    Dim reference As Double, fraction As Double, result As Double, seed As Variant
    mReverseTraceTest.ReadPoint mReverseTraceTest.Count, fraction, reference: seed = accepted
    If Not mReverseSolverTest.TryDeformationStep(mReverseBaseTest, mReverseDirectionTest, projection, coordinate, reference, result, _
            mReverseOptionsTest(3), mReverseOptionsTest(4), 1, consistentControl:=True, initialField:=seed) Then Err.Raise 5, , "Не подтверждён исходный успешный статус."
    mReverseSolverTest.DiscardStep
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\ArcFailureContext.Result.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMARCCONTEXT001"
    Put #handle, , magic
    Dim branch As CCoupledDeformationContinuation, mode As Long, outcome As Variant
    For mode = 0 To 2
        Set branch = New CCoupledDeformationContinuation
        outcome = branch.TestCachedArcFailureRecovery(mReverseSolverTest, projection, mode)
        If CBool(outcome(0)) <> (mode = 0) Then Err.Raise 5, , "Выбран отказ чужой проекции/ревизии либо потерян собственный отказ."
        TestWriteReverseBinaryVector handle, outcome
    Next mode
    Close #handle
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Выбор причины отказа изменил историю."
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
imports=['CCoupledDeformationContinuation','modTestCoupledElasticState']
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values.update({'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')',1).replace('\n','\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():
    if name.endswith(('.bas','.cls')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OriginalFailureContextAfterArc',actual_import_this_pass=imports,actual_component_count=256,
    arc_class_Newton_and_material_exact785=True,original_failure_requires_same_revision_and_exact_projection=True,
    directed_match_foreign_projection_foreign_revision_cases=True,no_new_physical_failure_rule=True,
    no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('786 SAME785 ARC/NEWTON/MATERIAL; STORED ORIGINAL FAILURE ONLY IN EXACT CONTEXT;3 DISPATCH CASES')
