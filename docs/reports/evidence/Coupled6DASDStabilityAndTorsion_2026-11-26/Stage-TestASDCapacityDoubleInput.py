"""ДЛЯ ТЕСТОВ. Повторяет тот же физический изгиб с вдвое большим исходным моментом и вдвое меньшим ожидаемым lambda, без перекомпиляции ядра."""
from pathlib import Path
from hashlib import sha256
import json,shutil,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1027ASDCapacityScaling_20261126';root=base/'Cold1028ASDCapacityDoubleInput_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
native=(parent/'Native/RealColdReplay.txt').read_text('utf-8-sig');assert 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in native
proof=json.loads((parent/'EarlySourceProof.json').read_bytes());assert proof['one_updated_driver_exact']
root.mkdir()
for directory in ('src','tests','tools','workbook','Input'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1','Run-TestASDBendingMainCapacity.ps1'):shutil.copyfile(parent/name,root/name)
shutil.copyfile(parent/'docs/regression/Diagnostic.xlsm',root/'docs/regression/Diagnostic.xlsm')
shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'docs/regression/CodeUpdate.json')
(root/'Input/CapacityBase.bin').write_bytes(struct.pack('<id',1,2000000.))
proof.update(book_reused_bitexact_from_parent=str(parent),all256_compiled_components_reused_exact=True,no_new_compilation_required=True)
(root/'EarlySourceProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=dict(directed_scope=dict(case='ASD_capacity_double_base_same_ray',base_Mx_Nmm=2000000.,
    exact_parent_book_reused=True,all_materials_and_steel_limits_and_equilibrium_tolerances_unchanged=True,
    physical_reference_steel_limit_Mx_kNm=4.658444129476362),
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
        if p.is_file() and p.relative_to(root).as_posix() not in ('docs/regression/Diagnostic.xlsm','docs/regression/CodeUpdate.json')})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1028; SAME ACTUAL COMPILED BOOK;BASE MX2M INSTEAD OF1M;SAME PHYSICAL RAY')
