"""ДЛЯ ТЕСТОВ. Готовит управление кривизной через существующий TryDeformationStep для проверки посттрещинной ветви чистого изгиба."""
from pathlib import Path
from hashlib import sha256
import json,shutil,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');parent=base/'Cold1009PureBendingASDOnly_20261126'
root=base/'Cold1011PureBendingASDCurvature_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
reference=json.loads((parent/'Input/PureBendingReferenceProof.json').read_bytes());assert reference['planar_reference_physical_limit_verified']
root.mkdir()
for directory in ('src','tests','tools','Input','workbook'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
files={'docs/regression/Diagnostic.xlsm':(parent/'docs/regression/Diagnostic.xlsm').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'CompiledSourceReuse.json':(json.dumps(dict(parent=str(parent)),indent=2)+'\n').encode(),
    'Guard-IsolatedOffice.ps1':(parent/'Guard-IsolatedOffice.ps1').read_bytes(),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'Run-TestPureBendingCurvature.ps1':(parent/'Run-TestPureBendingCoupled.ps1').read_bytes(),
    'TestPureBendingCurvature.bas.txt':(here/'TestPureBendingCurvature.bas.txt').read_bytes(),
    'Input/CurvatureOptions.bin':struct.pack('<i3d',3,2.,reference['Mx_ultimate_Nmm'],reference['deformation_plane'][1]),
    Path(__file__).name:Path(__file__).read_bytes()}
source=(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig');anchor='Public Function TestRunPureBendingCoupled(';assert source.count(anchor)==1
index=source.index(anchor);source=source[:index]+(here/'TestPureBendingCurvature.bas.txt').read_text('utf-8-sig')+'\n\n'+source[index:]
assert max(map(len,source.splitlines()))<1024
files['tests/modTestCoupledElasticState.bas']=source.replace('\n','\r\n').encode('utf-8-sig')
files['Run-TestPureBendingCurvature.ps1']=files['Run-TestPureBendingCurvature.ps1'].decode('utf-8-sig').replace('TestRunPureBendingCoupled','TestRunPureBendingCurvature').encode('utf-8-sig')
for name,data in files.items():(root/name).write_bytes(data)
prep=json.loads((parent/'Preparation.json').read_bytes());prep['inputs']={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
    if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'}
prep['directed_scope']=dict(case='PureBendingASDCurvatureControl',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=257,
    all256_numerical_components_exact1009=True,existing_TryDeformationStep=True,same_materials_and_force_tolerances=True,
    baseline_cold_load_fraction=.25,control='normalized kappaX in field degree2',physical_capacity_certified=False,
    no_ready_coupled_equilibrium_fields=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1011 CURVATURE CONTROL;EXISTING SOLVER;SAME MATERIALS/TOLERANCES;COLD BASELINE THEN NO FUTURE FIELDS')
