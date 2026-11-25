"""ДЛЯ ТЕСТОВ. Проверяет рабочее уточнение с явным бюджетом вызовов и защитой от несовпадающей истории; критерии не меняются."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold793AdaptivePreEvent_20261125';root=base/'Cold794PreEventBudget_20261125'
proof=json.loads((parent/'AdaptivePreEventProof.json').read_bytes());assert proof['returned_verified_candidate'] and len(proof['approaches'])==3
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PreEventBudgetCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
anchor='        error = -1#: lambdaError = -1#: lambdaBound = 0#: length = mSolver.SpatialFieldDistance(accepted, candidate)\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'        resolved = False: verified = False\n',1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Private Sub TestRunOperationalArc\(\).*?^End Sub',code)
anchor='    solved = arc.TryStep(mReverseFieldsTest, mReverseTraceTest, ratioData(1), 12, factor, projection)\n';assert method.count(anchor)==1
updated=method.replace(anchor,'''    inputHandle = FreeFile
    Open mPolishRootTest & "\\Input\\ArcCallAllowance.bin" For Binary Access Read As #inputHandle
    Dim allowanceData As Variant: allowanceData = TestReadPolishVector(inputHandle): Close #inputHandle
    TestArcMismatchedHistoryGuard arc
    solved = arc.TryStep(mReverseFieldsTest, mReverseTraceTest, ratioData(1), CLng(allowanceData(1)), factor, projection)
''',1)
updated=updated.replace('Or arc.Calls > 12','Or arc.Calls > CLng(allowanceData(1))',1)
code=code.replace(method,updated,1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отказ до решения для несовпадающего последнего поля; исходная коллекция и принятая материальная история не меняются.
Private Sub TestArcMismatchedHistoryGuard(ByVal arc As CCoupledArcContinuation)
    Dim fields As New Collection, field() As Double, before() As Double, after() As Double, i As Long
    For i = 1 To mReverseFieldsTest.Count
        field = mReverseFieldsTest(i)
        If i = mReverseFieldsTest.Count Then field(1) = field(1) + 0.0000000001
        fields.Add field
    Next i
    mReverseSolverTest.CopyAccepted before
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim result As Double, projection() As Double, solved As Boolean, number As Long, reason As String
    On Error Resume Next
    solved = arc.TryStep(fields, mReverseTraceTest, 1#, 12, result, projection)
    number = Err.Number: reason = Err.Description
    Err.Clear: On Error GoTo 0
    mReverseSolverTest.CopyAccepted after
    If number <> 5 Or solved Or arc.Calls <> 0 Or mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Не отклонена несовпадающая история до решения."
    For i = 1 To UBound(before): If before(i) <> after(i) Then Err.Raise 5, , "Проверка входа изменила принятое поле."
    Next i
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\ArcHistoryGuard.Result.txt" For Output As #handle
    Print #handle, "ERROR;" & CStr(number) & ";CALLS;" & CStr(arc.Calls) & ";REVISION_BEFORE;" & CStr(revision) & ";REVISION_AFTER;" & CStr(mReverseSolverTest.AcceptedRevision)
    Print #handle, "REASON;" & reason
    Close #handle
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
imports=['CCoupledArcContinuation','modTestCoupledElasticState'];compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values.update({'Input/ArcCallAllowance.bin':struct.pack('<id',1,12.),
    'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')',1).replace('\n','\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OperationalPreEventBudgetsAndInputGuard',actual_import_this_pass=imports,actual_component_count=256,
    source10_and_original_physics_exact793=True,per_level_flags_reset=True,explicit_call_allowance=12,
    mismatched_accepted_history_rejected_before_any_solver_call=True,
    no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('794 SAME793 ADAPTATION; PER-LEVEL FLAGS RESET; ALLOWANCE12 DATA; MISMATCHED HISTORY GUARD')
