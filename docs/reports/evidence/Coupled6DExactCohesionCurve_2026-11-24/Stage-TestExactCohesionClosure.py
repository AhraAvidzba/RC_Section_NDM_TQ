"""ДЛЯ ТЕСТОВ. Измеряет локальную работу и проверяет точное обновление координат при неизменных сцеплениях на полной кривой."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold708ExactCohesionClosure_20261124';parent=base/'Cold707SectionDilationClosure_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ExactCohesionCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
(root/'Matrix/Native').mkdir(parents=True,exist_ok=True)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='src/Materials/CCoupledConcreteResponse.cls';code=(root/path).read_text('utf-8-sig')
code=code.replace('Option Explicit\n','''Option Explicit
Private mExactCohesionUpdateTest As Boolean
Private mStateCallsTest As Long, mStateIterationsTest As Long, mClosureRequestsTest As Long, mClosureSuccessesTest As Long, mDilationReadsTest As Long
''',1)
method,=re.findall(r'(?ms)^Private Function EvaluatePointState\(.*?^End Function',code)
updated=method
anchor='    If Not mASDFlow Is Nothing Then\n';assert updated.count(anchor)==1
updated=updated.replace(anchor,'    mStateCallsTest = mStateCallsTest + 1\n'+anchor,1)
anchor='    Dim ct As Double, cc As Double, dt As Double, dc As Double, kt As Double, kc As Double\n';assert updated.count(anchor)==1
updated=updated.replace(anchor,anchor+'    Dim previousCtTest As Double, previousCcTest As Double\n',1)
anchor="    For j = 1 To MATERIAL_ITERATIONS + 1 ' После штатного цикла допускается один скалярный возврат.\n";assert updated.count(anchor)==1
updated=updated.replace(anchor,anchor+'        mStateIterationsTest = mStateIterationsTest + 1\n        previousCtTest = ct: previousCcTest = cc\n',1)
anchor='''        tension = tension + MAXIMUM_DAMPING * deltaTension
        compression = compression + MAXIMUM_DAMPING * deltaCompression''';assert updated.count(anchor)==1
updated=updated.replace(anchor,'''        If holdDilation And mExactCohesionUpdateTest And ct = previousCtTest And cc = previousCcTest Then
            tension = nextTension: compression = nextCompression ' При тех же сцеплениях следующий возврат даёт ту же пластическую поправку.
        Else
'''+anchor+'''
        End If''',1)
code=code.replace(method,updated,1)
anchor='        If NeedsDilationClosure(committed, updated, stress, virginElastic) Then\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'            mClosureRequestsTest = mClosureRequestsTest + 1\n',1)
anchor='            If TryDilationClosure(strain, committed, closed, closedStress) Then\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'                mClosureSuccessesTest = mClosureSuccessesTest + 1\n',1)
anchor='    TestReadDilationResidual = ';assert code.count(anchor)==0
anchor='    ReadDilationClosureResidual = EvaluatePointState';assert code.count(anchor)==1
code=code.replace(anchor,'    mDilationReadsTest = mDilationReadsTest + 1\n'+anchor,1)
code+='\n'+(here/'TestDilationWork.cls.txt').read_text('utf-8-sig')
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
path='tests/modTestCoupledElasticState.bas';code=(root/path).read_text('utf-8-sig')
anchor='        concrete(side).Initialize operators.PointCount, EC, 0.2, 16#, 30#, FT, 0.002, GT, 80#, LENGTH\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'        concrete(side).SetExactCohesionUpdateTest True\n',1)
anchor='    For Each strain In strainTargets\n';assert code.count(anchor)==1
code=code.replace(anchor,'    Dim progressStartedTest As Double: progressStartedTest = Timer\n'+anchor,1)
anchor='        solved = solver.TryDeformationStep(offset, direction, projection, CDbl(strain), predictionFactor, trialFactor, ABS_TOL, REL_TOL, 30, _';assert code.count(anchor)==1
code=code.replace(anchor,'''        If CDbl(strain) = 0.0008 Then TestCompareExactCohesionUpdate concrete(2), operators, seed, CDbl(strain), evidenceRoot, report, passed, failed
        TestWriteDilationProgress evidenceRoot, "BEGIN", CDbl(strain), progressStartedTest, concrete, solver.ResultMeta.ResultCode
'''+anchor,1)
anchor='            conditionHybridRecovery:=True, allowRefinedRestart:=True)\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'        TestWriteDilationProgress evidenceRoot, "END", CDbl(strain), progressStartedTest, concrete, solver.ResultMeta.ResultCode\n',1)
anchor='            TestOnsetNumber(expected) & ";" & TestOnsetNumber(concreteStress) & ";" & TestOnsetNumber(response(2).TrialSteelUtilization)\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''        If hasFile Then
            Close #handle
            Open evidenceRoot & "\\Native\\PostCrackTension.Curve.csv" For Append As #handle
        End If
''',1)
code+='\n'+(here/'TestDilationProgress.bas.txt').read_text('utf-8-sig')
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledConcreteResponse','modTestCoupledElasticState')")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
for name in (Path(__file__).name,'TestDilationWork.cls.txt','TestDilationProgress.bas.txt'):values[name]=(here/name).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ExactFixedCohesionUpdateWithWorkCounters',actual_import_this_pass=['CCoupledConcreteResponse','modTestCoupledElasticState'],
    original_material_parameters_and_acceptance=True,exact_equality_of_cohesions_required=True,
    local_original_vs_exact_same_history_comparison=True,expected_complete_curve_assertions=194,
    per_point_progress_flushed=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('EXACT COHESION UPDATE AND WORK COUNTERS STAGED; SAME ACCEPTANCE; FULL CURVE')
