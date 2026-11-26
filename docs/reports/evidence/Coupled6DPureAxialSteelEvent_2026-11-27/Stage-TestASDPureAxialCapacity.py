"""ДЛЯ ТЕСТОВ. Задаёт чистое растяжение или сжатие N в проверенной книге связанного ASD без изменения кода и допусков."""
from pathlib import Path
from hashlib import sha256
from argparse import ArgumentParser
import json, shutil, struct

parser = ArgumentParser(description=__doc__)
parser.add_argument('--case', type=int, required=True)
parser.add_argument('--sign', choices=('tension', 'compression'), required=True)
args = parser.parse_args()
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1037ASDNMTCapacity_20261126'
root = base / f'Cold{args.case}ASDPureN{args.sign.title()}_20261127'
assert not root.exists()
compiled = json.loads((parent / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
book = (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
assert sha256(book).hexdigest().upper() == compiled['WorkbookSHA256']
proof = json.loads((parent / 'EarlySourceProof.json').read_bytes())
assert proof['two_observation_components_exact']
root.mkdir()
for directory in ('src', 'tests', 'tools', 'workbook', 'Input'):
    shutil.copytree(parent / directory, root / directory)
for directory in ('Native', 'Temp', 'docs/regression'):
    (root / directory).mkdir(parents=True, exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1', 'Run-TestASDSpatialMainCapacity.ps1'):
    shutil.copyfile(parent / name, root / name)
(root / 'docs/regression/Diagnostic.xlsm').write_bytes(book)
shutil.copyfile(parent / 'docs/regression/CodeUpdate.json', root / 'docs/regression/CodeUpdate.json')
load = [100000. if args.sign == 'tension' else -100000., 0., 0., 0., 0., 0.]
(root / 'Input/CapacityLoad.bin').write_bytes(struct.pack('<i6di', 6, *load, 1))
proof.update(book_reused_bitexact_from_parent=str(parent), all256_compiled_components_reused_exact=True,
             no_new_compilation_required=True)
(root / 'EarlySourceProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
(root / 'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)), indent=2) + '\n', encoding='utf-8')
shutil.copyfile(Path(__file__), root / Path(__file__).name)
prep = dict(directed_scope=dict(case='ASD_pure_N_' + args.sign, target_N_Nmm=load,
            positive_N_means='tension', load_path='lambda*N', maximum_lambda=100.,
            compiled_book_exact_parent=True, all_materials_and_equilibrium_tolerances_unchanged=True,
            no_ready_equilibrium_input=True, physical_capacity_not_assumed=True),
            inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
                    for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix()
                    not in ('docs/regression/Diagnostic.xlsm', 'docs/regression/CodeUpdate.json')})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('PURE N PREPARED', root, load, 'EXPLICIT N PATH;ORIGINAL ASD,STEEL,TOLERANCES;NO NEW COMPILATION')
