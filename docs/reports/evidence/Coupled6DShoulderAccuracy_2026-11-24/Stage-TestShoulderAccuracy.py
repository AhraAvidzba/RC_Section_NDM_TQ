"""ДЛЯ ТЕСТОВ. Проверяет управляемое ужесточение равновесия плеч при сохранении исходной формулы погрешности и общего бюджета."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold717ShoulderAccuracy_20261124';parent=base/'Cold716ActualAdjointResidual_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ShoulderAccuracyCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='src/Solver/CCoupledDeformationContinuation.cls';code=(root/path).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Private Function Probe\(.*?^End Function',code);changed=method
anchor='        Optional ByVal robustDeparture As Boolean = False) As Boolean\n';assert changed.count(anchor)==1
changed=changed.replace(anchor,'''        Optional ByVal robustDeparture As Boolean = False, Optional ByVal accuracyScale As Double = 1#) As Boolean
    If accuracyScale <= 0# Or accuracyScale > 1# Then Err.Raise 5, "CCoupledDeformationContinuation", "Масштаб точности должен быть положительным и не превышать единицу."
    If mUsesEnvelopeControl And accuracyScale <> 1# Then Err.Raise 5, "CCoupledDeformationContinuation", "Для меры H уточнение задаётся отдельным алгоритмом."
''',1)
assert changed.count('point.Lambda, mAbsTol, mRelTol, mMaxIterations')==1
changed=changed.replace('point.Lambda, mAbsTol, mRelTol, mMaxIterations','point.Lambda, mAbsTol * accuracyScale, mRelTol * accuracyScale, mMaxIterations',1)
assert changed.count('mTrace, initialLambda, mAbsTol, mRelTol, _')==1
changed=changed.replace('mTrace, initialLambda, mAbsTol, mRelTol, _','mTrace, initialLambda, mAbsTol * accuracyScale, mRelTol * accuracyScale, _',1)
code=code.replace(method,changed,1)
method,=re.findall(r'(?ms)^Private Function ReplayShoulders\(.*?^End Function',code);changed=method
anchor='''            If Not mSolver.TrialLambdaUncertainty(mAbsTol, mRelTol, point.Uncertainty) Then
                mSolver.DiscardStep
                mProbeDetails = "Не удалось оценить чувствительность множителя к невязке."
                Exit Function
            End If''';assert changed.count(anchor)==1
changed=changed.replace(anchor,'''            If Not ConfirmShoulderAccuracy(point, previousLambda, samples(index).Coordinate, keepRight) Then
                mSolver.DiscardStep
                Exit Function
            End If''',1)
code=code.replace(method,changed,1)
code+='\n'+(here/'CandidateShoulderAccuracy.cls.txt').read_text('utf-8-sig')
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledDeformationContinuation'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
for name in (Path(__file__).name,'CandidateShoulderAccuracy.cls.txt'):values[name]=(here/name).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='GoalDirectedShoulderAccuracy',actual_import_this_pass=['CCoupledDeformationContinuation'],
    original_adjoint_uncertainty_formula_exact=True,force_tolerance_only_tightened_at_critical_points=True,
    all_quality_retries_count_as_ordinary_probes=True,original96_probe_limit=True,original30_primary_iterations=True,
    original_capacity_accuracy=.001,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('717 CRITICAL POINT ACCURACY; ORIGINAL ADJOINT BOUND; ALL RETRIES INSIDE96 PROBES')
