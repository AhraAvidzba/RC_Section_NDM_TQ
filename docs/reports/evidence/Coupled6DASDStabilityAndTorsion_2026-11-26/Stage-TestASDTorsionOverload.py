"""ДЛЯ ТЕСТОВ. Проверяет определяющий хомут и статус отказа при исходном кручении4кНм тем же исполняемым кодом."""
from pathlib import Path
from hashlib import sha256
import json,shutil,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1035ASDPureTorsion_20261126';root=base/'Cold1036ASDTorsionOverload_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
book=(parent/'docs/regression/Diagnostic.xlsm').read_bytes();assert sha256(book).hexdigest().upper()==compiled['WorkbookSHA256']
verified=json.loads((parent/'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes());assert verified['controlled_component']==4 and verified['all56plus4_equations_and_six_resultants_verified']
proof=json.loads((parent/'EarlySourceProof.json').read_bytes());assert proof['one_updated_driver_exact']
root.mkdir()
for directory in ('src','tests','tools','workbook','Input'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1','Run-TestASDBendingMainCapacity.ps1'):shutil.copyfile(parent/name,root/name)
(root/'docs/regression/Diagnostic.xlsm').write_bytes(book)
shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'docs/regression/CodeUpdate.json')
(root/'Input/CapacityBase.bin').write_bytes(struct.pack('<id',1,4000000.))
assert (root/'Input/CapacityComponent.bin').read_bytes()==struct.pack('<ii',1,4)
proof.update(book_reused_bitexact_from_parent=str(parent),all256_compiled_components_reused_exact=True,no_new_compilation_required=True)
(root/'EarlySourceProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=dict(directed_scope=dict(case='ASD_pure_torsion_overload',base_T_Nmm=4000000.,load_path='lambda*T',maximum_lambda=100.,
    actual_compiled_book_reused_bitexact=True,original_materials_and_equilibrium_tolerances_unchanged=True,
    expected_physical_event_not_forced=True,check_hoop_event_and_status_independently=True),
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
        if p.is_file() and p.relative_to(root).as_posix() not in ('docs/regression/Diagnostic.xlsm','docs/regression/CodeUpdate.json')})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1036; PURE T4KNM;SAME COMPILED BOOK;HOOP STATUS CHECK;NO ACCEPTED FIELD INPUT')
