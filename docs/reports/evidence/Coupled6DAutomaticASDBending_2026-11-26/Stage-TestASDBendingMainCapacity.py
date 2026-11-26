"""ДЛЯ ТЕСТОВ. Готовит поиск предела через штатный рабочий CCoupledSectionAnalysis вместо отдельной кривой с ручным выбором меры."""
from pathlib import Path
from hashlib import sha256
import shutil,json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');parent=base/'Cold1018SelectableMaterialCore_20261126'
root=base/'Cold1020ASDBendingMainCapacity_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
proof=json.loads((parent/'SelectableMaterialCoreProof.json').read_bytes());assert proof['all_eight_working_core_cases_completed']
assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
root.mkdir()
for directory in ('src','tests','tools','workbook'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1','Compile-TestSelectableMaterialCore.ps1'):shutil.copyfile(parent/name,root/name)
shutil.copyfile(parent/'docs/regression/Diagnostic.xlsm',root/'docs/regression/Diagnostic.xlsm')
shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'ParentCodeUpdate.json')
source=(parent/'tests/modTestSelectableMaterialCore.bas').read_text('utf-8-sig')+'\n'+(here/'TestASDBendingMainCapacity.bas.txt').read_text('utf-8-sig')
(root/'tests/modTestSelectableMaterialCore.bas').write_bytes(source.replace('\n','\r\n').encode('utf-8-sig'))
runner=(parent/'Run-TestSelectableMaterialCore.ps1').read_text('utf-8-sig').replace('modTestSelectableMaterialCore.TestRunSelectableMaterialCore','modTestSelectableMaterialCore.TestRunASDBendingMainCapacity')
(root/'Run-TestASDBendingMainCapacity.ps1').write_bytes(runner.replace('\n','\r\n').encode('utf-8-sig'))
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=dict(directed_scope=dict(case='WorkingASDBendingMainCapacity',actual_import_this_pass=['modTestSelectableMaterialCore'],
    all255_other_components_unchanged=True,physical_capacity_not_assumed=True,ASD_priority=True,
    exact_steel_limits=[-.018,.025],postyield_tangent_MPa=1500.,base_Mx_Nmm=1000000.,reference_ASD_strain_limit_kNm=4.658444129476362),
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
        if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1020; WORKING ASD SolveCapacity;SAME MATERIALS/STEEL AS VERIFIED BENDING STRAIN LIMIT;NO READY FIELDS')
