"""ДЛЯ ТЕСТОВ. Отдельно исследует чувствительность к численному допуску; рабочие допуски и исходные контрольные ожидания не меняет."""
from pathlib import Path
from hashlib import sha256
import json,shutil

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1030ASDScale1M_20261126';root=base/'Cold1034ASDTolerance1E5_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
reference=json.loads((parent/'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes());assert reference['all56plus4_equations_and_six_resultants_verified']
root.mkdir()
for directory in ('src','tests','tools','workbook','Input'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1','Run-TestASDBendingMainCapacity.ps1'):shutil.copyfile(parent/name,root/name)
shutil.copyfile(parent/'docs/regression/Diagnostic.xlsm',root/'docs/regression/Diagnostic.xlsm')
shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'ParentCodeUpdate.json')
name='src/Solver/CCoupledSectionAnalysis.cls';code=(root/name).read_text('utf-8-sig')
for key in ('STRESS_TOLERANCE','RELATIVE_TOLERANCE'):
    needle=f'Private Const {key} As Double = 0.0000005';assert code.count(needle)==1
    code=code.replace(needle,f'Private Const {key} As Double = 0.00001',1)
(root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
name='tests/modTestSelectableMaterialCore.bas';code=(root/name).read_text('utf-8-sig')
needle='0.001, 0.0000005, 0.0000005, mMainCapacityBaseTest)';assert code.count(needle)==1
code=code.replace(needle,'0.001, 0.00001, 0.00001, mMainCapacityBaseTest)',1)
(root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
(root/'Input/NumericalTolerances.json').write_text(json.dumps(dict(absolute_MPa=.00001,relative=.00001,
    reference_absolute_MPa=.0000005,reference_relative=.0000005,study_only=True,production_tolerance_unchanged=True),sort_keys=True,indent=2)+'\n',encoding='utf-8')
compiler=(base/'Cold1027ASDCapacityScaling_20261126/Compile-TestASDCapacityScaling.ps1').read_text('utf-8-sig')
compiler=compiler.replace("-ModuleNames 'modTestSelectableMaterialCore'","-ModuleNames @('CCoupledSectionAnalysis','modTestSelectableMaterialCore')")
(root/'Compile-TestASDToleranceStudy.ps1').write_bytes(compiler.replace('\n','\r\n').encode('utf-8-sig'))
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=dict(directed_scope=dict(case='ASD_numerical_tolerance_sensitivity',base_Mx_Nmm=1000000.,
    study_absolute_tolerance_MPa=.00001,study_relative_tolerance=.00001,original_control_expectations_unchanged=True,
    existing_hybrid_gate_preserved_disabled_for_ASD=True,physical_materials_and_steel_limits_unchanged=True,
    independent_balance_errors_must_be_reported_in_N_and_Nmm=True,production_parameters_unchanged=True),
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
        if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1034; PRIVATE TOLERANCE SENSITIVITY1E-5;NO MATERIAL CHANGES;ORIGINAL CONTROLS AND PRODUCTION TOLERANCES RETAINED')
