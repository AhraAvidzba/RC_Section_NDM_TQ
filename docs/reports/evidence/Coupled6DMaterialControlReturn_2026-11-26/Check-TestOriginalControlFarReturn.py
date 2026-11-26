"""ДЛЯ ТЕСТОВ. Независимо проверяет прямой возврат942 и полную цену материального прогноза938 плюс возврата."""
from pathlib import Path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold942OriginalControlFarReturn_20261126';parent=base/'Cold941OriginalControlReturn_20261126'
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
script=here/'Check-TestOriginalControlReturn.py';code=script.read_text('utf-8-sig')
# Вложенный общий проверяющий936 тоже должен читать942, поэтому меняются оба точных вхождения имени941.
assert code.count('Cold941OriginalControlReturn_20261126')==2
code=code.replace('Cold941OriginalControlReturn_20261126','Cold942OriginalControlFarReturn_20261126')
scope={'__file__':str(script),'__name__':'original_control_far_return_test'}
exec(compile(code,str(script),'exec'),scope)
proof=json.loads((root/'OriginalControlReturnProof.json').read_bytes())
proof.pop('initial_native940_material_root_residual_bitexact_by_guard')
proof.update(initial_native938_material_root_residual_bitexact_by_guard=True,
    immutable_compiled941_reused=True,intermediate940_root_not_used=True)
starting=json.loads((root/'VerifiedStartingMaterialRoot.json').read_bytes())
for case in proof['cases']:
    case['combined_material938_and_original_return_steps']=starting['steps']+case['steps']
    case['combined_material938_and_original_return_calls']=starting['calls']+case['calls']
    case['original_failed_primary_attempt_effort_included']=False
(root/'OriginalControlReturnProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
