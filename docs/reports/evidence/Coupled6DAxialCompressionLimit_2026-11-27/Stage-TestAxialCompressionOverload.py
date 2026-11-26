"""ДЛЯ ТЕСТОВ. Проверяет отказ при сжатии 1000 кН на той же собранной книге, меняя только осевую нагрузку."""
from pathlib import Path
from hashlib import sha256
import json, shutil, struct

base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1056ASDPureNFullCompressionBranch_20261127'
root = base / 'Cold1058ASDPureNCompressionOverload_20261127'
assert not root.exists()
execution = json.loads((parent / 'MainCapacityExecutionProof.json').read_bytes())
assert execution['native']['own_process_closed'] and execution['has_ultimate_point'] == 1
for replay in ('Native/OpenSeesMainCapacityEquilibrium.json',
               'FullCompressionBranchTestReplay/Native/OpenSeesMainCapacityEquilibrium.json'):
    assert json.loads((parent / replay).read_bytes())['all56plus4_equations_and_six_resultants_verified']
compiled = json.loads((parent / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
book = (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
assert sha256(book).hexdigest().upper() == compiled['WorkbookSHA256']
proof = json.loads((parent / 'EarlySourceProof.json').read_bytes())
assert proof['changed_components_exact'] and proof['controller_change_only'] == 'full-branch'
root.mkdir()
for directory in ('src', 'tests', 'tools', 'workbook', 'Input'):
    shutil.copytree(parent / directory, root / directory)
for directory in ('Native', 'Temp', 'docs/regression'):
    (root / directory).mkdir(parents=True, exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1', 'Run-TestASDSpatialMainCapacity.ps1'):
    shutil.copyfile(parent / name, root / name)
(root / 'docs/regression/Diagnostic.xlsm').write_bytes(book)
shutil.copyfile(parent / 'docs/regression/CodeUpdate.json', root / 'docs/regression/CodeUpdate.json')
load = [ -1000000., 0., 0., 0., 0., 0. ]
(root / 'Input/CapacityLoad.bin').write_bytes(struct.pack('<i6di', 6, *load, 1))
proof.update(book_reused_bitexact_from_parent=str(parent), no_new_compilation_required=True,
             all256_compiled_components_reused_exact=True)
(root / 'EarlySourceProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
(root / 'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)), indent=2) + '\n', encoding='utf-8')
shutil.copyfile(Path(__file__), root / Path(__file__).name)
prep = dict(parent=str(parent), directed_scope=dict(case='ASD_pure_N_compression_overload',
            target_N_Nmm=load, load_path='lambda*N', positive_N_means='tension',
            compiled_book_exact_parent=True, only_load_changed=True,
            zero_initial_material_history=True, all_materials_and_tolerances_unchanged=True,
            physical_failure_not_assumed=True),
            inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
                    for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix()
                    not in ('docs/regression/Diagnostic.xlsm', 'docs/regression/CodeUpdate.json')})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('COMPRESSION OVERLOAD PREPARED', root, 'ONLY N CHANGED;SAME COMPILED BOOK;COLD START')
