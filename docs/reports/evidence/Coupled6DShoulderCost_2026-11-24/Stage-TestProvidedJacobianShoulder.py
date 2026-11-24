"""ДЛЯ ТЕСТОВ. Сравнивает существующий гибридный резерв с собранной касательной вместо полной разностной матрицы."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold730ProvidedJacobianShoulder_20261124';parent=base/'Cold728FastShoulderWitness_20261124'
assert json.loads((parent/'ShoulderWitnessProof.json').read_bytes())['quality_confirmed']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ProvidedJacobianCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='RunTestPostCrackRawReplay = RunTestCoupledPostCrackTension(evidenceRoot, False, False, 0, True)';assert code.count(anchor)==1
code=code.replace(anchor,'RunTestPostCrackRawReplay = RunTestCoupledPostCrackTension(evidenceRoot, False, False, 1, True)',1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=(here/'Compile-TestExternalPoint.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ExistingProvidedJacobianHybridShoulder',actual_import_this_pass=['modTestCoupledElasticState'],
    existing_SetPostCrackTangentModeTest=1,only_hybrid_Jacobian_selection_changed=True,
    exact728_material_and_solvers=True,local_tolerances_unchanged=True,original52prefix_and_target_unchanged=True,
    same30_iterations_three_quality_attempts_and_global_tolerance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('730 EXISTING HYBRID PROVIDED JACOBIAN; EXACT728 PHYSICS/PREFIX/QUALITY; NO NEW CORE ALGORITHM')
