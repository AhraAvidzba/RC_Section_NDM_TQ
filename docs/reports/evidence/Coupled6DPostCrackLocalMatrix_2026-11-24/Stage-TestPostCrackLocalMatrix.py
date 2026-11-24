"""ДЛЯ ТЕСТОВ. Сравнивает готовую и уточнённую материальные касательные на полной посттрещинной кривой из нуля."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json


def TestStage(root,parent):
    """Не меняет закон материала, нагрузки, допуски и30 итераций; проверка кривой отделена от длинного Trace."""
    here=Path(__file__).resolve().parent
    run_path(str(here/'Stage-TestRealPostCrack.py'))['TestStage'](root,parent,False,True,False,False,False)
    prep=json.loads((root/'Preparation.json').read_bytes());values={}
    path='src/Solver/CCoupledNewtonSolver.cls';code=(root/path).read_text('utf-8-sig')
    anchor='Option Explicit\n';assert code.count(anchor)==1
    code=code.replace(anchor,anchor+'\nPrivate mPostCrackTangentModeTest As Long\n',1)
    before='''    converged = hybrid.TrySolve(equations, coordinates, ordinaryCalls, ordinarySteps, absoluteTolerance, _
        0.00000000001, 0.2, solution, relativeDifferenceStep:=HYBRID_DIFFERENCE_STEP, _
        automaticScaling:=conditionRecovery, freshJacobian:=conditionRecovery, captureDifferenceBracket:=(reserved > 0), _
        differenceObserver:=observer, completeNewtonLineSearch:=(reserved > 0))'''
    assert code.count(before)==1
    code=code.replace(before,'''    If conditionRecovery And referenceLambda = 0# Then
        Dim differenceStepTest As Double: differenceStepTest = HYBRID_DIFFERENCE_STEP
        If mPostCrackTangentModeTest > 0 Then differenceStepTest = 0#
        converged = hybrid.TrySolveLeastSquaresTrust(equations, coordinates, ordinaryCalls, ordinarySteps, _
            absoluteTolerance, solution, differenceStepTest)
    Else
'''+before+'''
    End If''',1)
    before='''    If Not Evaluate(field, withJacobian, recordTrial:=False, consistentBranch:=refinedBranch, refinedBranch:=refinedBranch) Then Exit Function'''
    assert code.count(before)==1
    code=code.replace(before,'''    Dim useRefinedTest As Boolean: useRefinedTest = refinedBranch Or (withJacobian And mPostCrackTangentModeTest = 2)
    If Not Evaluate(field, withJacobian, recordTrial:=False, consistentBranch:=useRefinedTest, refinedBranch:=useRefinedTest) Then Exit Function''',1)
    code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выбирает только численный источник матрицы; ноль сохраняет полную разностную матрицу,1 — собранную,2 — уточнённую по той же ветви.
Friend Sub SetPostCrackTangentModeTest(ByVal mode As Long)
    If mode < 0 Or mode > 2 Or mHybridActive Then Err.Raise 5, "CCoupledNewtonSolver", "Некорректный режим тестовой касательной."
    mPostCrackTangentModeTest = mode
End Sub
'''
    values[path]=code
    path='tests/modTestCoupledElasticState.bas';code=(root/path).read_text('utf-8-sig')
    before='Public Function RunTestCoupledPostCrackTension(Optional ByVal evidenceRoot As String = "", Optional ByVal capacityOnly As Boolean = False) As String'
    assert code.count(before)==1
    code=code.replace(before,before.replace(') As String',', Optional ByVal prescribedOnly As Boolean = False, Optional ByVal tangentModeTest As Long = 0) As String'),1)
    before='    Dim solver As New CCoupledNewtonSolver: solver.Initialize operators, response(2), EC, response(3), response(1)\n'
    assert code.count(before)==1;code=code.replace(before,before+'    solver.SetPostCrackTangentModeTest tangentModeTest\n',1)
    before='    If Not solver.ReplayAcceptedFields(history, history.Count) Then Err.Raise 5, , "Не восстановлено начало до первой трещины."'
    assert code.count(before)==1;code=code.replace(before,'    If prescribedOnly Then GoTo Done\n'+before,1)
    before='report & IIf(capacityOnly, "TOTAL_COUPLED_POSTCRACK_CAPACITY", "TOTAL_COUPLED_POSTCRACK_TENSION")'
    assert code.count(before)==1
    code=code.replace(before,'report & IIf(prescribedOnly, "TOTAL_COUPLED_POSTCRACK_PRESCRIBED", IIf(capacityOnly, "TOTAL_COUPLED_POSTCRACK_CAPACITY", "TOTAL_COUPLED_POSTCRACK_TENSION"))',1)
    for mode,label in ((1,'Matrix'),(2,'Refined')):
        code+=f'''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет полную предписанную кривую от нулевой истории с режимом касательной{mode}; поиск предела выполняется отдельно.
Public Function RunTestPostCrack{label}(ByVal evidenceRoot As String) As String
    RunTestPostCrack{label} = RunTestCoupledPostCrackTension(evidenceRoot, False, True, {mode})
End Function
'''
        (root/label/'Native').mkdir(parents=True)
        runner=(root/'Run-TestRealPostCrack.ps1').read_text('utf-8-sig')
        runner=runner.replace('modTestCoupledElasticState.RunTestCoupledPostCrackTension','modTestCoupledElasticState.RunTestPostCrack'+label)
        runner=runner.replace('PostCrackTension.txt','Curve'+label+'.txt').replace('-MacroArgument1 $Root','-MacroArgument1 (Join-Path $Root '+repr(label)+')')
        values['Run-TestPostCrack'+label+'.ps1']=runner
    values[path]=code
    for name,code in values.items():
        raw=code.replace('\n','\r\n').encode('utf-8-sig');(root/name).write_bytes(raw);prep['inputs'][name]=sha256(raw).hexdigest().upper()
    name=Path(__file__).name;(root/name).write_bytes(Path(__file__).read_bytes());prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
    prep['directed_scope'].update(case='PostCrackLocalMatrixFromZero',prescribed_curve_only=True,
       nominal_targets=21,no_controller_trace=True,max_primary_iterations_per_point=30,expected_assertions_per_complete_curve=192,
       assembled_matrix_modes=[1,2],material_and_final_residual_unchanged=True)
    (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('POSTCRACK LOCAL MATRIX CURVES STAGED; FROM ZERO; NO TRACE; TWO SOURCES OF J')


if __name__=='__main__':
    base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
    TestStage(base/'Cold700PostCrackLocalMatrix_20261124',base/'Cold699WorkingRankKernel_20261124')
