"""ДЛЯ ТЕСТОВ. Публикует только проверенную сохранную книгу и её фактический полный VBA-дамп с контролем исходных хешей."""
from pathlib import Path
from hashlib import sha256
import json, os

repo = Path.cwd()
root = Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1074WorkingASDPreservedOutput_20261127')
proof = json.loads((root / 'WorkingASDPreservedOutputProof.json').read_bytes())
assert proof['all121_non_vba_payloads_bitexact'] and proof['actual_full_dump_equal_verified_compiled_candidate']
assert proof['native_foundations']['failed'] == 0 and proof['native_foundations']['macros'] == 7
pairs = [('RC_Section_NDM.xlsm', 'saved_candidate_sha256'), ('VBA_All_Code.txt', 'actual_dump_sha256')]
before, after, payloads, originals, temporaries = {}, {}, {}, {}, {}
for name, key in pairs:
    path = repo / 'workbook/output' / name
    data = (root / 'docs/regression' / ('Diagnostic.xlsm' if name.endswith('.xlsm') else name)).read_bytes()
    assert sha256(data).hexdigest().upper() == proof[key]
    original = (root / 'Before' / name).read_bytes()
    assert path.read_bytes() == original, 'Public file changed; preserve new user data before publishing: ' + name
    before[name] = sha256(original).hexdigest().upper()
    after[name] = sha256(data).hexdigest().upper()
    payloads[name], originals[name] = data, original
    temp = repo / 'tmp/material-consistency' / ('PublishWorkingASD.' + name)
    assert not temp.exists()
    temp.write_bytes(data)
    temporaries[name] = temp
replaced = []
try:
    for name, key in pairs:
        path = repo / 'workbook/output' / name
        assert path.read_bytes() == originals[name]
        os.replace(temporaries[name], path)
        replaced.append(name)
    for name, key in pairs:
        assert (repo / 'workbook/output' / name).read_bytes() == payloads[name]
except BaseException:
    for name in reversed(replaced):
        path = repo / 'workbook/output' / name
        assert path.read_bytes() == payloads[name], 'Concurrent public change; automatic rollback forbidden.'
        temp = repo / 'tmp/material-consistency' / ('RollbackWorkingASD.' + name)
        temp.write_bytes(originals[name])
        os.replace(temp, path)
    raise
publication = dict(before=before, after=after, source_user_data_preserved_bitexact=True,
    actual_dump_from_reopened_book=True, default_material_law_not_changed=True,
    Config_ASD_selector_pending=True, no_full_acceptance_claim=True, main_not_merged=True,
    whole_goal_completed=False)
(root / 'OutputPublication.json').write_text(json.dumps(publication, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('ACTUAL OUTPUT AND FULL CODE DUMP UPDATED', json.dumps(publication, sort_keys=True))
