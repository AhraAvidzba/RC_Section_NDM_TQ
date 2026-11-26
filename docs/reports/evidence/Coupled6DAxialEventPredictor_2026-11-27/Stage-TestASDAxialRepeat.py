"""ДЛЯ ТЕСТОВ. Повторяет осевой поиск на холодном старте и при другом масштабе исходного N без изменения исполняемого кода."""
from pathlib import Path
from hashlib import sha256
from argparse import ArgumentParser
import json, shutil, struct

parser = ArgumentParser(description=__doc__)
parser.add_argument('--variant', choices=('repeat', 'scale'), required=True)
variant = parser.parse_args().variant
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1042ASDPureNBudget_20261127'
name, force = {'repeat': ('Cold1044ASDPureNColdRepeat_20261127', 100000.),
               'scale': ('Cold1045ASDPureNScale10kN_20261127', 10000.)}[variant]
root = base / name
assert not root.exists()
compiled = json.loads((parent / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
book = (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
assert sha256(book).hexdigest().upper() == compiled['WorkbookSHA256']
assert json.loads((parent / 'MainCapacityExecutionProof.json').read_bytes())['native']['own_process_closed']
proof = json.loads((parent / 'EarlySourceProof.json').read_bytes())
assert proof['changed_components_exact'] and proof['controller_change_only'] == 'budget'
root.mkdir()
for directory in ('src', 'tests', 'tools', 'workbook', 'Input'):
    shutil.copytree(parent / directory, root / directory)
for directory in ('Native', 'Temp', 'docs/regression'):
    (root / directory).mkdir(parents=True, exist_ok=True)
for filename in ('Guard-IsolatedOffice.ps1', 'Run-TestASDSpatialMainCapacity.ps1'):
    shutil.copyfile(parent / filename, root / filename)
(root / 'docs/regression/Diagnostic.xlsm').write_bytes(book)
shutil.copyfile(parent / 'docs/regression/CodeUpdate.json', root / 'docs/regression/CodeUpdate.json')
(root / 'Input/CapacityLoad.bin').write_bytes(struct.pack('<i6di', 6, force, 0., 0., 0., 0., 0., 1))
proof.update(book_reused_bitexact_from_parent=str(parent), all256_compiled_components_reused_exact=True,
             no_new_compilation_required=True)
(root / 'EarlySourceProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
(root / 'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)), indent=2) + '\n', encoding='utf-8')
shutil.copyfile(Path(__file__), root / Path(__file__).name)
prep = dict(directed_scope=dict(case='ASD_pure_N_' + variant, target_N_Nmm=[force, 0., 0., 0., 0., 0.],
            positive_N_means='tension', load_path='lambda*N', maximum_lambda=100.,
            compiled_book_exact_parent=True, all_materials_and_equilibrium_tolerances_unchanged=True,
            zero_initial_material_history=True, physical_capacity_not_assumed=True),
            inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
                    for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix()
                    not in ('docs/regression/Diagnostic.xlsm', 'docs/regression/CodeUpdate.json')})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('PURE N REPEAT PREPARED', root, 'N0=', force, 'EXACT192-PROBE KERNEL;COLD START')
