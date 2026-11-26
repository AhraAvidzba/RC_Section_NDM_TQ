"""ДЛЯ ТЕСТОВ. Готовит независимые холодные пробы чистого Mx в текущем CDP и тестовом ASD относительно плоского предела."""
from pathlib import Path
from hashlib import sha256
import json,shutil,struct

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1005PureBendingReference_20261126';root=base/'Cold1006PureBendingCoupled_20261126';assert not root.exists()
proof=json.loads((parent/'PureBendingReferenceProof.json').read_bytes());assert proof['planar_reference_physical_limit_verified']
root.mkdir()
for directory in ('src','tests','tools','Input','workbook'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
files={
    'docs/regression/Diagnostic.xlsm':(parent/'docs/regression/Diagnostic.xlsm').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'CompiledSourceReuse.json':(json.dumps(dict(parent=str(parent)),indent=2)+'\n').encode(),
    'Guard-IsolatedOffice.ps1':(parent/'Guard-IsolatedOffice.ps1').read_bytes(),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'Run-TestPureBendingCoupled.ps1':(parent/'Run-TestPureBendingReference.ps1').read_bytes(),
    'Input/MxReference.bin':struct.pack('<id',1,proof['Mx_ultimate_Nmm']),
    'Input/PureBendingReferenceProof.json':(parent/'PureBendingReferenceProof.json').read_bytes(),
    'TestPureBendingCoupled.bas.txt':(here/'TestPureBendingCoupled.bas.txt').read_bytes(),
    Path(__file__).name:Path(__file__).read_bytes(),
}
code=(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
anchor='Public Function TestRunPureBendingPlanarReference(';assert code.count(anchor)==1
index=code.index(anchor);code=code[:index]+(here/'TestPureBendingCoupled.bas.txt').read_text('utf-8-sig')+'\n\n'+code[index:]
assert max(map(len,code.splitlines()))<1024
files['tests/modTestCoupledElasticState.bas']=code.replace('\n','\r\n').encode('utf-8-sig')
files['Run-TestPureBendingCoupled.ps1']=files['Run-TestPureBendingCoupled.ps1'].decode('utf-8-sig').replace('TestRunPureBendingPlanarReference','TestRunPureBendingCoupled').encode('utf-8-sig')
contract=json.loads((parent/'Input/MaterialContract.json').read_bytes());contract['section_target']=[0.,proof['Mx_ultimate_Nmm'],0.,0.,0.,0.]
contract['control_purpose']='cold pure-bending probes around stock planar physical ultimate; coupled core explicitly forced'
files['Input/MaterialContract.json']=(json.dumps(contract,sort_keys=True,indent=2)+'\n').encode()
files['Input/PureBendingModels.json']=(json.dumps(dict(reference_Mx_Nmm=proof['Mx_ultimate_Nmm'],ratios=[.25,.5,.75,.9,.95,.98,1.,1.02,1.05,1.1,1.2],
    each_case_cold=True,force_coupled_kernel_despite_T_Q_zero=True,
    model1=dict(law='Barcelona-CDP',E_MPa=30000.,poisson=.2,fc0_MPa=16.,fc_MPa=30.,ft_MPa=3.,peak_compression_strain=.002,Gt_N_per_mm=.12,Gc_N_per_mm=80.,localization_mm=125.),
    model2=dict(law='ASD native single-update test',tables='matched.csv',stress_floors='same as primary OpenSees source'),
    steel=dict(E_MPa=200000.,long_fy_tension_MPa=400.,long_fy_compression_MPa=350.,hoop_fy_MPa=240.,postyield_tangent_MPa=1500.,long_tension_strain_limit=.025,long_compression_strain_limit=.018,hoop_strain_limit=.02),
    differences_from_planar=['spatial concrete has tensile cracking/softening; planar strength ignores tensile concrete','spatial steel hardens; planar TwoLine steel has a plateau','transverse reinforcement and transverse strains are present in spatial model'],
    no_ready_coupled_equilibrium_fields=True,numerical_failure_is_not_physical_failure=True),sort_keys=True,indent=2)+'\n').encode()
for name,data in files.items():(root/name).write_bytes(data)
prep=json.loads((parent/'Preparation.json').read_bytes())
prep['inputs']={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
    if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'}
prep['directed_scope']=dict(case='PureBendingColdCoupledProbes',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=257,
    all256_numerical_components_exact1005=True,no_ready_coupled_equilibrium_input=True,target_count=11,material_model_count=2,
    original_force_tolerances=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1006;22 COLD PURE-MX PROBES;CURRENT CDP AND ASD NATIVE UPDATE;STOCK PLANAR REFERENCE ONLY')
