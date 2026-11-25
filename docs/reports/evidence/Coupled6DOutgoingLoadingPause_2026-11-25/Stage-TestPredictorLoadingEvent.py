"""ДЛЯ ТЕСТОВ. Разрешает искать разрешимый переход от исходного прогноза, когда обе равновесные дуговые пробы не сошлись."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold804HistoryAccuracyEndpoint_20261125';root=base/'Cold807PredictorLoadingEvent_20261125'
proof=json.loads((base/'Cold806EndpointPointWitness_20261125/EndpointPointWitnessProof.json').read_bytes())
assert proof['native']['own_process_closed'] and proof['predictor_brackets_real_loading_transition']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PredictorEventCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());name='src/Solver/CCoupledArcContinuation.cls'
code=(root/name).read_text('utf-8-sig')
old='''    ElseIf (sf Or sh) And mCalls + 3 <= callAllowance Then
        Dim bracket() As Double, bracketLambda As Double
        If sf Then bracket = full: bracketLambda = lf Else bracket = half: bracketLambda = lh
'''
new='''    ElseIf mCalls + 3 <= callAllowance Then
        Dim bracket() As Double, bracketLambda As Double
        If sf Then
            bracket = full: bracketLambda = lf
        ElseIf sh Then
            bracket = half: bracketLambda = lh
        Else
            bracket = fullSeed: bracketLambda = fullPrediction ' Прогноз только задаёт отрезок поиска поверхности; равновесие и пригодность найденного события проверяются заново.
        End If
'''
assert code.count(old)==1;code=code.replace(old,new,1)
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values={name:code.replace('\n','\r\n').encode('utf-8-sig'),
 'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",
 "-ModuleNames 'CCoupledArcContinuation'",1).replace('\n','\r\n').encode('utf-8-sig'),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='LoadingEventFromPredictorAfterBothArcRootFailures',
 actual_import_this_pass=['CCoupledArcContinuation'],actual_component_count=256,
 Newton_controller_materials_and_test_module_exact804=True,all45source_fields_exact804=True,
 predictor_is_bracket_only_never_accepted=True,first_resolved_loading_event_required=True,
 event_equilibrium_half_field_lambda_and_resolution_checks_unchanged=True,
 exact_restore_required=True,pair30_event30_and_call_allowance12_unchanged=True,
 physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('807 PREDICTOR BRACKET IF BOTH ARC ROOTS FAIL; SAME804 PHYSICS/ACCEPTANCE/BUDGETS')
