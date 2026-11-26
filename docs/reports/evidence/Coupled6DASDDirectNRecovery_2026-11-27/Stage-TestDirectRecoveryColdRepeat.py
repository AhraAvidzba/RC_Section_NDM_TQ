"""ДЛЯ ТЕСТОВ. Повторяет шесть осевых состояний ASD на холодном старте с побайтно той же книгой и всеми прежними входами."""
from pathlib import Path
from hashlib import sha256
import json, shutil

base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1065ASDPureNDirectEarlyPrefix_20261127'
root = base / 'Cold1066ASDPureNDirectColdRepeat_20261127'
assert not root.exists()
observed = json.loads((parent / 'ASDDirectPureNExecutionProof.json').read_bytes())
assert observed['native']['own_process_closed'] and observed['all56plus4_equations_and_six_resultants_verified']
assert [c['status'] for c in observed['cases']] == ['OK', 'OK', 'OK', 'FAIL', 'OK', 'FAIL']
compiled = json.loads((parent / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
book = (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
assert sha256(book).hexdigest().upper() == compiled['WorkbookSHA256']
for directory in ('src', 'tests', 'tools', 'Input', 'workbook'):
    shutil.copytree(parent / directory, root / directory)
for directory in ('Native', 'Temp', 'docs/regression'):
    (root / directory).mkdir(parents=True, exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1', 'Run-TestASDDirectPureN.ps1', 'docs/regression/CodeUpdate.json'):
    shutil.copyfile(parent / name, root / name)
(root / 'docs/regression/Diagnostic.xlsm').write_bytes(book)
source = json.loads((parent / 'EarlySourceProof.json').read_bytes())
source.update(compiled_book_reused_bitexact_from_parent=str(parent), no_new_compilation_required=True,
              all256_compiled_components_reused_exact=True)
(root / 'EarlySourceProof.json').write_text(json.dumps(source, sort_keys=True, indent=2) + '\n', encoding='utf-8')
for index in range(1, 7):
    case = root / f'Case{index:02d}'
    shutil.copytree(parent / f'Case{index:02d}/Input', case / 'Input')
    for directory in ('Native', 'docs/regression'):
        (case / directory).mkdir(parents=True, exist_ok=True)
    shutil.copyfile(root / 'EarlySourceProof.json', case / 'EarlySourceProof.json')
    (case / 'docs/regression/Diagnostic.xlsm').write_bytes(book)
(root / 'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent),
    only_cold_native_execution=True, all256_components_bitexact=True, no_compilation=True), indent=2) + '\n', encoding='utf-8')
shutil.copyfile(Path(__file__), root / Path(__file__).name)
prep = dict(parent=str(parent), directed_scope=dict(case='ASDDirectPureNColdRepeat',
    no_input_or_code_or_material_change=True, six_cases_start_independent_zero_histories=True,
    actual_saved_book_reused_bitexact=True, no_CDP_run=True),
    inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
            for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix()
            not in ('docs/regression/Diagnostic.xlsm', 'docs/regression/CodeUpdate.json')})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('ASD DIRECT NDS COLD REPEAT PREPARED;EXACT SAME BOOK,INPUTS AND MATERIALS;NO NEW COMPILATION')
