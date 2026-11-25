"""ДЛЯ ТЕСТОВ. Направляет однозначный недавно принятый переход сразу в проверяемую материальную координату с общим бюджетом трёх вызовов."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold819AutomaticLoadingArc_20261125';root=base/'Cold821PriorityLoadingArc_20261125'
assert json.loads((parent/'AutomaticLoadingArcProof.json').read_bytes())['pair']['verified']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PriorityLoadingCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
anchor='    prediction = halfSeed: mCalls = mCalls + 1';assert code.count(anchor)==1
code=code.replace(anchor,'''    Dim recent As Variant, loadingDeparture As Boolean
    loadingDeparture = mSolver.TryReadRecentLoadingDeparture(loadingPrediction, recent)
    Dim fullPrediction As Double: fullPrediction = factor(4) + ratio * dl
    If Not loadingDeparture Then
    prediction = halfSeed: mCalls = mCalls + 1''',1)
anchor='    Dim fullPrediction As Double: fullPrediction = factor(4) + ratio * dl\n    If sh Then';assert code.count(anchor)==1
code=code.replace(anchor,'    If sh Then',1)
anchor='    mSolver.DiscardStep\n    Dim verified As Boolean';assert code.count(anchor)==1
code=code.replace(anchor,'''    mSolver.DiscardStep
    Else
        ReDim half(1 To 1): ReDim full(1 To 1): ReDim ah(1 To 1): ReDim af(1 To 1)
    End If
    Dim verified As Boolean''',1)
anchor='        Dim recent As Variant\n        If mSolver.TryReadRecentLoadingDeparture(loadingPrediction, recent) Then';assert code.count(anchor)==1
code=code.replace(anchor,'        If loadingDeparture Then',1)
# Блоки имеют явные уровни отступа; алгоритмы и условия внутри них не меняются.
start=code.index('    If Not loadingDeparture Then\n')+len('    If Not loadingDeparture Then\n')
end=code.index('    Else\n        ReDim half(1 To 1)',start)
code=code[:start]+''.join('    '+line if line.strip() else line for line in code[start:end].splitlines(True))+code[end:]
start=code.index('        Else\n        Dim bracket()')+len('        Else\n')
end=code.index('        End If\n    End If\n    Dim restored',start)
code=code[:start]+''.join('    '+line if line.strip() else line for line in code[start:end].splitlines(True))+code[end:]
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='    TestRunOperationalArc\n    mPolishExecutedTest = True';assert code.count(anchor)==1
code=code.replace(anchor,'    TestRecentLoadingDirectionGuards\n    TestRunOperationalArc\n    TestRecentLoadingResetGuard\n    mPolishExecutedTest = True',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Нулевой прогноз и разгрузка не разрешают исходящую координату; читающие пробы сохраняют принятую ревизию.
Private Sub TestRecentLoadingDirectionGuards()
    Dim accepted() As Double, old() As Double, prediction() As Double, info As Variant, i As Long
    accepted = mReverseFieldsTest(mReverseFieldsTest.Count): old = mReverseFieldsTest(mReverseFieldsTest.Count - 1)
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    If mReverseSolverTest.TryReadRecentLoadingDeparture(accepted, info) Then Err.Raise 5, , "Нулевой прогноз ошибочно разрешил исходящую координату."
    prediction = accepted
    For i = 1 To mReverseSolverTest.PhysicalDegreeCount: prediction(i) = old(i): Next i
    If mReverseSolverTest.TryReadRecentLoadingDeparture(prediction, info) Then Err.Raise 5, , "Разгрузка ошибочно разрешила исходящую координату."
    For i = 1 To mReverseSolverTest.PhysicalDegreeCount: prediction(i) = 2# * accepted(i) - old(i): Next i
    If Not mReverseSolverTest.TryReadRecentLoadingDeparture(prediction, info) Then Err.Raise 5, , "Положительный прогноз потерял принятую границу."
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Чтение границы изменило принятую ревизию."
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\RecentLoading.Guards.txt" For Output As #handle
    Print #handle, "ZERO_REJECTED;UNLOADING_REJECTED;FORWARD_SELECTED;REVISION_UNCHANGED"
    Close #handle
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сброс независимой истории аннулирует границу; поле из прежней ревизии не создаёт новый контекст.
Private Sub TestRecentLoadingResetGuard()
    Dim old() As Double: old = mReverseFieldsTest(mReverseFieldsTest.Count)
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    mReverseSolverTest.ResetHistory
    Dim info As Variant
    If mReverseSolverTest.TryReadRecentLoadingDeparture(old, info) Then Err.Raise 5, , "Сброшенная история сохранила исходящую координату."
    If mReverseSolverTest.AcceptedRevision <= revision Then Err.Raise 5, , "Сброс не изменил ревизию."
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\RecentLoading.Guards.txt" For Append As #handle
    Print #handle, "RESET_CONTEXT_REJECTED;REVISION_ADVANCED"
    Close #handle
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');imports=['CCoupledArcContinuation','modTestCoupledElasticState']
values.update({'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')',1).replace('\n','\r\n').encode('utf-8-sig'),
    'Input/ArcCallAllowance.bin':struct.pack('<id',1,3.),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='PriorityRecentLoadingArcWithinThreeCalls',actual_import_this_pass=imports,actual_component_count=256,
    immutable_material_and_Newton_sources_exact819=True,automatic_recent_coordinate_before_arc_pair=True,
    no_change_to_any_pair_physics_or_acceptance_criterion=True,call_allowance3=True,
    direction_and_reset_negative_controls=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
