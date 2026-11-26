"""ДЛЯ ТЕСТОВ. Готовит парное сравнение сечения НДМ и объёмного OpenSees при нулевых Q для отделения градиента моментов вдоль призмы."""
from pathlib import Path
from hashlib import sha256
import json,shutil

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1002ASDExactFloors_20261126';root=base/'Cold1003ASDNMTCold_20261126'
volume=base/'Cold1004OpenSeesNMTCore_20261126'
assert not root.exists() and not volume.exists()
prior=json.loads((parent/'ASDHardShearColdProof.json').read_bytes())
assert prior['OpenSees_material_tensor_equivalence_verified'] and prior['all_six_resultants_and56plus4_equations_verified']
root.mkdir()
for directory in ('src','tests','tools','Input','workbook'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Diagnostic.xlsm',):shutil.copyfile(parent/'docs/regression'/name,root/'docs/regression'/name)
files={
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'CompiledSourceReuse.json':(json.dumps(dict(parent=str(parent)),indent=2)+'\n').encode(),
    'Guard-IsolatedOffice.ps1':(parent/'Guard-IsolatedOffice.ps1').read_bytes(),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'Run-TestASDHardShearCold.ps1':(parent/'Run-TestASDHardShearCold.ps1').read_bytes(),
    'Replay-TestASDNativeMaterial.py':(here/'Replay-TestASDNativeMaterial.py').read_bytes(),
    Path(__file__).name:Path(__file__).read_bytes(),
}
source=(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
start=source.index('Public Function TestRunASDHardShearCold(');end=source.index('Public Function TestStartFixedForceComparison(',start)
body=source[start:end]
anchor='    Dim loads As Variant: loads = TestHardShearLoads()';assert body.count(anchor)==1
body=body.replace(anchor,'    Dim loads As Variant: loads = Array(28000#, 510000#, 180000#, 1500000#, 0#, 0#) \' Нулевые Q убирают градиент M вдоль объёмного контроля.',1)
old='marks = Array(0.5, 0.65, 0.699036407470703, 1#, 1.10521946491058, 1.2)'
assert body.count(old)==1
body=body.replace(old,'marks = Array(0.5, 0.65, 0.699036407470703, 0.936482644081116, 0.974817657470703, 1#, 1.10263192386604, 1.10521946491058, 1.2)',1)
body=body.replace('Public Function TestRunASDHardShearCold(','Public Function TestRunASDNMTCold(')
body=body.replace('TestRunASDHardShearCold =','TestRunASDNMTCold =')
files['tests/modTestCoupledElasticState.bas']=(source[:start]+body+source[end:]).replace('\n','\r\n').encode('utf-8-sig')
files['Run-TestASDHardShearCold.ps1']=files['Run-TestASDHardShearCold.ps1'].decode('utf-8-sig').replace('TestRunASDHardShearCold','TestRunASDNMTCold').encode('utf-8-sig')
target=[28000.,510000.,180000.,1500000.,0.,0.]
contract=json.loads((parent/'Input/MaterialContract.json').read_bytes())
contract.update(section_target=target,control_purpose='isolate moment gradients from transverse forces; all six-force production requirements retained')
files['Input/MaterialContract.json']=(json.dumps(contract,indent=2,sort_keys=True)+'\n').encode()
for name,data in files.items():(root/name).write_bytes(data)
prep=json.loads((parent/'Preparation.json').read_bytes())
prep['inputs']={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
                if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'}
prep['directed_scope']=dict(case='ColdNMTWithoutMomentGradient',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=257,
    all256_numerical_components_exact1002=True,no_ready_equilibrium_input=True,virgin_material_history=True,
    original_material_tables_steel_and_force_tolerances=True,only_Qx_Qy_zero=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
volume.mkdir()
for name,data in [('Guard-IsolatedOffice.ps1',files['Guard-IsolatedOffice.ps1']),
                  ('validate_opensees_hardshear.py',(repo/'tools/validate_opensees_hardshear.py').read_bytes())]:(volume/name).write_bytes(data)
vprep=dict(variant='matched',section_target=target,length_mm=800.,cut_mm=400.,dz_mm=25.,nonlinear_length_mm=125.,
    end_condition='balanced-tractions',capture_cut=True,requested_factors=[.5,.65,.6990364074707031,.9364826440811157,.9748176574707031,1.,1.10263192386604,1.1052194649105769,1.2],
    no_moment_gradient_from_Q=True,auxiliary_elastic_load_transfer_regions=True,homogeneous_beam_equivalence=False,
    source_sha256=sha256((volume/'validate_opensees_hardshear.py').read_bytes()).hexdigest().upper(),physical_capacity_certified=False)
(volume/'Preparation.json').write_text(json.dumps(vprep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('PREPARED1003/1004 NMT;Q=0;SAME MATERIAL TABLES AND STEEL;NO FUTURE EQUILIBRIUM INPUT')
