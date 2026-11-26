"""ДЛЯ ТЕСТОВ. Повторяет объёмный изгиб с точной строковой передачей Mx вместо округляющего числового аргумента PowerShell."""
from pathlib import Path
from hashlib import sha256
import json

base=Path('D:/CodexScratch/RC_Section_NDM_TQ');parent=base/'Cold1008PureBendingVolume_20261126';root=base/'Cold1010PureBendingVolumeExactInput_20261126';assert not root.exists()
before=json.loads((parent/'Result.json').read_bytes());prep=json.loads((parent/'Preparation.json').read_bytes())
assert before['termination'] is not None and before['section_target']!=prep['section_target']
root.mkdir()
for name in ('Guard-IsolatedOffice.ps1','validate_opensees_hardshear.py','PureBendingReferenceProof.json'):(root/name).write_bytes((parent/name).read_bytes())
prep.update(exact_cli_transfer='reference double formatted R with invariant culture and passed as a string',
    preceding_attempt='1008',preceding_Mx_input_difference_Nmm=before['section_target'][1]-prep['section_target'][1])
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
(root/Path(__file__).name).write_bytes(Path(__file__).read_bytes())
print('PREPARED1010 EXACT INPUT;PRIOR ONE-ULP CLI DIFFERENCE',prep['preceding_Mx_input_difference_Nmm'],'NMM;NO MODEL OR TOLERANCE CHANGED')
