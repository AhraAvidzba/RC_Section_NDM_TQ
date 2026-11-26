"""ДЛЯ ТЕСТОВ. Проверяет короткие доли демпфирования возле поверхности нагружения21/+1 при прежних30 итерациях."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold968PrefixNewtonCapture_20261126';root=base/'Cold969PrefixSpatialFloor_20261126'
proof=json.loads((parent/'PrefixNewtonCaptureProof.json').read_bytes())
assert proof['returned_binary_bitexact967'] and len(proof['rejected_newton_records'])==3
assert all(r['probes'][-1]['fraction']==1/128 for r in proof['rejected_newton_records'])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'PrefixSpatialFloorCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
updates={};name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
old='Private Const MIN_FRACTION As Double = 0.0078125';assert code.count(old)==1
constant="Private Const SPATIAL_MIN_FRACTION As Double = 0.0000152587890625 ' Управляемые пространственные пробы разрешают короткое пересечение материальной границы."
code=code.replace(old,old+'\n'+constant,1)
old='        If mSteelBranchBuilds > 0 Then minimumFraction = STEEL_MIN_FRACTION';assert code.count(old)==1
code=code.replace(old,old+'\n        If mControlActive And mUseSpatialHistory Then minimumFraction = SPATIAL_MIN_FRACTION',1)
begin=code.index('Private Function TryProspectiveBranchStep(');end=code.index('\nEnd Function',begin)
block=code[begin:end];old='    fraction = 1#\n    Do While fraction >= MIN_FRACTION';assert block.count(old)==1
new='''    Dim minimumFraction As Double: minimumFraction = MIN_FRACTION
    If mControlActive And mUseSpatialHistory Then minimumFraction = SPATIAL_MIN_FRACTION
    fraction = 1#
    Do While fraction >= minimumFraction'''
code=code[:begin]+block.replace(old,new,1)+code[end:]
updates[name]=code.replace('\n','\r\n').encode('utf-8-sig')
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='OriginalScalarPrefixWithSharedSpatialDampingFloor1Over65536',actual_import_this_pass=['CCoupledNewtonSolver'],
    all256_other_actual_components_exact968=True,only_SolveStep_and_TryProspectiveBranchStep_floor_changed=True,
    shared_primary_and_prospective_spatial_floor=1/65536,additional_Newton_iterations_not_granted=True,
    primary30_root10_and_original_equilibrium_half_steel_criteria_unchanged=True,all_material_laws_exact968=True,
    actual_rejected_directions_recorded=True,no_release_promotion=True,physical_capacity_certified=False)
prep['directed_scope'].pop('all257_actual_components_and_saved_book_bitexact967',None)
prep['directed_scope']['no_new_numerical_code_or_compilation']=False
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'SAME30/ROOT10;SPATIAL PRIMARY+PROSPECTIVE FLOOR1/65536;MATERIAL AND ACCEPTANCE UNCHANGED')
