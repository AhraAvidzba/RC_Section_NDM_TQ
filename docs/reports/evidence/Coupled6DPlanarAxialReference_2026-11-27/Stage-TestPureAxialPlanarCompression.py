"""ДЛЯ ТЕСТОВ. Меняет только исходное N в проверенной книге осевого плоского контроля для холодного расчёта сжатия."""
from pathlib import Path
from hashlib import sha256
import json, shutil, struct

base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1059PureAxialPlanarTension_20261127'
root = base / 'Cold1060PureAxialPlanarCompression_20261127'
assert not root.exists()
proof = json.loads((parent / 'PureAxialPlanarReferenceProof.json').read_bytes())
assert proof['native']['own_process_closed'] and proof['planar_reference_physical_limit_verified']
book = (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
source = json.loads((parent / 'PlanarAxialSourceProof.json').read_bytes())
assert sha256(book).hexdigest().upper() == source['saved_book_sha256']
for directory in ('src', 'tests', 'tools', 'Input', 'workbook'):
    shutil.copytree(parent / directory, root / directory)
for directory in ('Native', 'Temp', 'docs/regression'):
    (root / directory).mkdir(parents=True, exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1', 'Run-TestPureAxialPlanarReference.ps1', 'CurrentPlanarKernelProof.json',
             'PlanarAxialSourceProof.json', 'docs/regression/CodeUpdate.json'):
    shutil.copyfile(parent / name, root / name)
(root / 'docs/regression/Diagnostic.xlsm').write_bytes(book)
(root / 'Input/PlanarAxialBase.bin').write_bytes(struct.pack('<id', 1, -100000.))
control = json.loads((root / 'Input/PlanarAxialControl.json').read_bytes())
control['base_N_N'] = -100000.
(root / 'Input/PlanarAxialControl.json').write_text(json.dumps(control, sort_keys=True, indent=2) + '\n', encoding='utf-8')
(root / 'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent), compiled_book_reused_bitexact=True,
    only_N_changed=True, zero_initial_history=True), indent=2) + '\n', encoding='utf-8')
shutil.copyfile(Path(__file__), root / Path(__file__).name)
prep = dict(parent=str(parent), directed_scope=dict(case='PureAxialPlanarCompression',
            base_N_N=-100000., no_new_compilation=True, actual_planar_kernel_exact_parent=True,
            no_artificial_diagram_extension=True, physical_capacity_not_assumed=True),
            inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
                    for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix()
                    not in ('docs/regression/Diagnostic.xlsm', 'docs/regression/CodeUpdate.json')})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('PLANAR PURE N COMPRESSION PREPARED;SAME COMPILED BOOK;ONLY N CHANGED')
