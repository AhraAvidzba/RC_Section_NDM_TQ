"""ДЛЯ ТЕСТОВ. Проверяет прогноз дилатансии по принятой негидростатической ветви вместо почти осевого дроблёного прогноза."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold710PriorDilationSeed_20261124';parent=base/'Cold709PostYieldLocalProbe_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PriorDilationCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
(root/'Matrix/Native').mkdir(parents=True,exist_ok=True)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='src/Materials/CCoupledConcreteResponse.cls';code=(root/path).read_text('utf-8-sig')
anchor='    Dim center As Double: center = CurrentDilation(guide.TensionEnergy, guide.CompressionEnergy, r, trace / damageFactor, q / damageFactor)\n'
assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''    Dim priorTrace As Double, priorQ As Double, priorHigh As Double, priorMiddle As Double, priorLow As Double
    Dim priorFraction As Double, priorDamage As Double
    CoupledStressInvariants committed.Stress, priorTrace, priorQ, priorHigh, priorMiddle, priorLow
    priorFraction = PositiveStressFraction(priorHigh, priorMiddle, priorLow)
    priorDamage = (1# - priorFraction * committed.TensionDamage) * (1# - committed.CompressionDamage)
    If priorDamage > 0# Then
        If priorQ > TENSILE_AXIS_BAND_RATIO * mFc * priorDamage + PLASTIC_ROUNDOFF * GeomMax(Abs(priorHigh), Abs(priorLow)) Then
            center = CurrentDilation(committed.TensionEnergy, committed.CompressionEnergy, priorFraction, priorTrace / priorDamage, priorQ / priorDamage) ' Прогноз по последней принятой ветви; полный шаг заново проверяет текущий закон.
        End If
    End If
''',1)
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledConcreteResponse'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='PriorAcceptedDilationSeed',actual_import_this_pass=['CCoupledConcreteResponse'],
    original_full_step_equations_and_acceptance=True,only_dilation_initial_center_changed=True,
    current_law_and_yield_rechecked=True,future_analytical_answer_not_used=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('710 PRIOR ACCEPTED DILATION PREDICTOR; SAME FULL STEP EQUATIONS AND FINAL ACCEPTANCE')
