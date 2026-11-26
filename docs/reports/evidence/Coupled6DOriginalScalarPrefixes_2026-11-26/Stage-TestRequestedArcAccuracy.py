"""ДЛЯ ТЕСТОВ. Сравнивает автоматическую дугу на исходной заданной точности с внутренним ужесточением1e-4; настройки физики неизменны."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold963IncomingEventBudget_20261126';root=base/'Cold964RequestedArcAccuracy_20261126'
proof=json.loads((parent/'IncomingEventBudgetProof.json').read_bytes());assert not proof['returned_verified_candidate']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'RequestedArcAccuracyCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
updates={};name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
old='Private Const ACCURACY_SCALE As Double = 0.0001';assert code.count(old)==1
code=code.replace(old,'Private Const ACCURACY_SCALE As Double = 1#',1)
old="' Сохраняет подготовленное ядро и исходную траекторию; внутренние пробы используют более строгую точность без изменения материала."
new="' Сохраняет подготовленное ядро и исходную траекторию; внутренние пробы используют исходную заданную точность без изменения материала."
assert code.count(old)==1;code=code.replace(old,new,1)
updates[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledArcContinuation'",1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OriginalRequestedArcAccuracyInsteadOfInternal1eMinus4',actual_import_this_pass=['CCoupledArcContinuation'],
    actual_component_count=257,all256_other_components_exact963=True,only_accuracy_scale_and_comment_changed=True,
    source_user_options_and_force_units_bitexact963=True,internal_accuracy_scale_from=.0001,internal_accuracy_scale_to=1.,
    internal_equilibrium_and_lambda_noise_and_resolution_follow_original_requested_tolerance=True,
    field_error_limit005_materials_positive_lambda_orientation_locality_and30_phase96_calls_unchanged=True,
    no_future_guide_or_field=True,no_material_history_acceptance_inside_arc=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'INTERNAL ACCURACY SCALE1;ORIGINAL REQUESTED OPTIONS UNCHANGED;ALL ACTUAL RESIDUALS REQUIRE INDEPENDENT CHECK')
