"""ДЛЯ ТЕСТОВ. Усиливает только внутреннее решение полных/половинных возвратов958, сохраняя исходные критерии приёмки и неопределённости."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold958BoundedIncomingSteps_20261126';root=base/'Cold959AccurateIncomingSteps_20261126'
proof=json.loads((parent/'BoundedIncomingStepsProof.json').read_bytes())
assert proof['accepted_steps']==1 and not proof['reached_current_material_event']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'AccurateIncomingCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
updates={};name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
anchor='Private Const YIELD_EVENT_HALF_ERROR As Double = 0.005';assert code.count(anchor)==1
start=code.index(anchor);end=code.index('\n',start)
code=code[:end]+"\nPrivate Const MATERIAL_RETURN_SOLVE_FRACTION As Double = 0.0001 ' Уточняет только внутренние равновесия возврата; приёмка формы, сил и оценка неопределённости используют исходную точность."+code[end:]
for method,next_method in [('Private Function TryReturnOriginalMaterialControl','Private Function ConfirmOriginalReturnHalf'),
 ('Private Function ConfirmOriginalReturnHalf','Private Function ICoupledDifferenceObserver_AcceptBracket')]:
    begin=code.index(method);end=code.index(next_method,begin);body=code[begin:end]
    old='        absoluteTolerance, relativeTolerance, remaining, consistentControl:='
    assert body.count(old)==1,(method,body.count(old))
    body=body.replace(old,'        absoluteTolerance * MATERIAL_RETURN_SOLVE_FRACTION, relativeTolerance * MATERIAL_RETURN_SOLVE_FRACTION, remaining, consistentControl:=',1)
    code=code[:begin]+body+code[end:]
updates[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledNewtonSolver'",1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith('.cls'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='AccurateInternalFullAndHalfIncomingEquilibria',actual_import_this_pass=['CCoupledNewtonSolver'],
    actual_component_count=257,all256_other_components_exact958=True,internal_solve_tolerance_factor=.0001,
    original_force_acceptance_half_shape_lambda_noise_uncertainty_and_steel_preserved=True,
    guide_and_all_rejected_attempts_share30_per_accepted_step=True,source_first_guide_cost2_counted=True,
    test_entry_and_input_bitexact958=True,earlier_primary_and_predictor_effort_not_included=True,
    original25_restoration_required=True,no_production_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'INTERNAL FULL/HALF SOLVE1e-4;ALL ORIGINAL ACCEPTANCE AND UNCERTAINTY GUARDS RETAINED')
