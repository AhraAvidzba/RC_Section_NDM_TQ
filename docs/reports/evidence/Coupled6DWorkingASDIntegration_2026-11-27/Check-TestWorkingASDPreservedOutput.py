"""ДЛЯ ТЕСТОВ. Принимает сохранный VBA-кандидат только по точным частям пакета, фактическому коду, повторному открытию и направленным нативным тестам."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from zipfile import ZipFile
import json, re

repo = Path.cwd()
here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
root = base / 'Cold1074WorkingASDPreservedOutput_20261127'
source = json.loads((root / 'PreservedPackageProof.json').read_bytes())
parent = Path(source['parent'])
candidate = root / 'docs/regression/Diagnostic.xlsm'
assert sha256(candidate.read_bytes()).hexdigest().upper() == source['saved_candidate_sha256']
with ZipFile(root / 'Before/RC_Section_NDM.xlsm') as old, ZipFile(candidate) as new, ZipFile(parent / 'docs/regression/Diagnostic.xlsm') as compiled:
    assert old.namelist() == new.namelist() and old.comment == new.comment
    assert new.read('xl/vbaProject.bin') == compiled.read('xl/vbaProject.bin')
    for name, digest in source['non_vba_parts'].items():
        assert old.read(name) == new.read(name)
        assert sha256(new.read(name)).hexdigest().upper() == digest
assert len(source['non_vba_parts']) == 121
inspection = json.loads((parent / 'OutputInspection.json').read_bytes())
assert not inspection['numerical_reference_differences'] and not inspection['dump_differences']
assert [item['name'] for item in inspection['source_checks'] if not item['exact_current_source']] == ['CCoupledElasticState', 'CCrackFormationCalculator']
critical = {'CCoupledDeformationContinuation', 'CCoupledLoadContinuation', 'CCoupledNewtonSolver', 'CCoupledSectionAnalysis'}
assert critical.issubset({item['name'] for item in inspection['source_checks'] if item['exact_current_source']})
reopen = json.loads((root / 'Native/PreservedReopenProof.json').read_text('utf-8-sig'))
assert reopen['SavedReopened'] and reopen['Compiled'] and reopen['ReadOnly'] and reopen['NoSave'] and reopen['Components'] == 255
assert json.loads((root / 'Native/ReopenClosure.json').read_text('utf-8-sig'))['State'] == 'Absent'
native, text = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'Foundations')
assert native['macros'] == 7 and native['failed'] == 0 and not native['timeout600'] and native['own_process_closed']


def TestDump(path):
    """Читает полный нативный CodeModule-дамп, сохраняя все строки и литералы."""
    lines = path.read_text('utf-8-sig').splitlines()
    result, index = {}, 0
    while index < len(lines):
        if not lines[index].startswith('COMPONENT: '):
            index += 1
            continue
        name = lines[index][11:]
        assert name not in result
        count = int(lines[index + 2][7:])
        result[name] = '\n'.join(lines[index + 4:index + 4 + count]).strip('\r\n')
        index += 4 + count
    return result


actual_dump = TestDump(root / 'docs/regression/VBA_All_Code.txt')
compiled_dump = TestDump(parent / 'docs/regression/VBA_All_Code.txt')
assert len(actual_dump) == len(compiled_dump) == 255 and actual_dump == compiled_dump
proof = dict(saved_candidate_sha256=source['saved_candidate_sha256'],
    actual_dump_sha256=sha256((root / 'docs/regression/VBA_All_Code.txt').read_bytes()).hexdigest().upper(),
    actual_compiled_vba_sha256=source['actual_compiled_vba_sha256'], actual_components=255,
    all121_non_vba_payloads_bitexact=True, names_cells_styles_shapes_validations_and_package_order_preserved=True,
    actual_full_dump_equal_verified_compiled_candidate=True, saved_reopened_native_compilation=True,
    all_numerical_vba_matches_verified_1067=True, private_diagnostic_module_not_in_output=True,
    native_foundations=native, source_display_exceptions=['CCoupledElasticState author comment transliteration',
        'CCrackFormationCalculator finite-Double guard displayed with 15 significant digits by VBE'],
    no_new_CDP_run=True, no_full_acceptance_run=True, dropdown_UI_not_claimed=True,
    default_material_law_not_changed=True, Config_ASD_choice_UI_pending=True,
    whole_user_run_10_seconds_not_met=True, public_output_unchanged_at_check=True, whole_goal_completed=False)
(root / 'WorkingASDPreservedOutputProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('PRESERVED WORKING ASD OUTPUT VERIFIED', json.dumps(proof, sort_keys=True))
