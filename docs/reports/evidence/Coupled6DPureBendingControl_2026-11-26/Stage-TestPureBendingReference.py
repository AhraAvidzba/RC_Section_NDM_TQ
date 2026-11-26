"""ДЛЯ ТЕСТОВ. Готовит штатный плоский предел для следующего контроля чистого изгиба; рабочая книга и численные классы неизменны."""
from pathlib import Path
from hashlib import sha256
import json,shutil

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1003ASDNMTCold_20261126';root=base/'Cold1005PureBendingReference_20261126';assert not root.exists()
assert json.loads((parent/'ASDHardShearColdProof.json').read_bytes())['OpenSees_material_tensor_equivalence_verified']
root.mkdir()
for directory in ('src','tests','tools','Input','workbook'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
files={
    'docs/regression/Diagnostic.xlsm':(parent/'docs/regression/Diagnostic.xlsm').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'CompiledSourceReuse.json':(json.dumps(dict(parent=str(parent)),indent=2)+'\n').encode(),
    'Guard-IsolatedOffice.ps1':(parent/'Guard-IsolatedOffice.ps1').read_bytes(),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'Run-TestPureBendingReference.ps1':(parent/'Run-TestASDHardShearCold.ps1').read_bytes(),
    Path(__file__).name:Path(__file__).read_bytes(),
    'TestPureBendingPlanarReference.bas.txt':(here/'TestPureBendingPlanarReference.bas.txt').read_bytes(),
}
code=(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
anchor='Public Function TestRunASDNMTCold(';assert code.count(anchor)==1
index=code.index(anchor);code=code[:index]+(here/'TestPureBendingPlanarReference.bas.txt').read_text('utf-8-sig')+'\n\n'+code[index:]
assert max(map(len,code.splitlines()))<1024
files['tests/modTestCoupledElasticState.bas']=code.replace('\n','\r\n').encode('utf-8-sig')
files['Run-TestPureBendingReference.ps1']=files['Run-TestPureBendingReference.ps1'].decode('utf-8-sig').replace('TestRunASDNMTCold','TestRunPureBendingPlanarReference').encode('utf-8-sig')
files['Input/PureBendingControl.json']=(json.dumps(dict(width_mm=180.,height_mm=120.,longitudinal_bars=[dict(x_mm=0.,y_mm=y,diameter_mm=12.) for y in (-30.,30.)],
    concrete_diagram='ThreeLine',concrete_tension='Ignore',steel_diagram='TwoLine',diagram_extension=False,
    reference_method='CCapacitySolver.SolveByUltimateStrain',reference_path='lambda*Mx',reference_base_Mx_Nmm=1000000.,
    selected_coupled_targets_relative_to_planar_limit=[.25,.5,.75,.9,.95,.98,1.,1.02,1.05,1.1,1.2],
    full_user_request='Compare planar ultimate and equilibrium availability in explicit coupled kernel and external engineering solvers; distinguish material-model differences from numerical failure',
    physical_capacity_of_coupled_model_not_assumed=True),sort_keys=True,indent=2)+'\n').encode()
for name,data in files.items():(root/name).write_bytes(data)
prep=json.loads((parent/'Preparation.json').read_bytes())
prep['inputs']={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
    if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'}
prep['directed_scope']=dict(case='PureBendingPlanarReference',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=257,
    all256_numerical_components_exact1003=True,only_new_test_planar_reference=True,artificial_diagram_extension=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1005 STOCK PLANAR ULTIMATE Mx;NO DIAGRAM EXTENSION;SAME SECTION;256 NUMERICAL COMPONENTS UNCHANGED')
