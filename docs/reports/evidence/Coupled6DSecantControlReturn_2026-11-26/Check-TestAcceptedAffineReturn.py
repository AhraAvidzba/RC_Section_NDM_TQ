"""ДЛЯ ТЕСТОВ. Проверяет возврат946 из секущей с принятым состоянием и записывает суммарную цену материального прогноза."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;script=here/'Check-TestAffineControlReturn.py';code=script.read_text('utf-8-sig')
assert code.count('Cold944AffineControlReturn_20261126')==2
code=code.replace('Cold944AffineControlReturn_20261126','Cold946AcceptedAffineReturn_20261126')
scope={'__file__':str(script),'__name__':'accepted_affine_return_check_test'};exec(compile(code,str(script),'exec'),scope)
root=scope['root'];proof=json.loads((root/'OriginalControlReturnProof.json').read_bytes())
proof.pop('initial_native943_secant_residual_bitexact_by_guard')
proof.update(initial_native945_secant_residual_bitexact_by_guard=True,low_material936_solve_not_used=True,
    low_endpoint_is_accepted25=True)
for case in proof['cases']:
    case['material938_high_root_and_return_steps']=17+case['steps']
    case['earlier_failed_primary_effort_included']=False
(root/'OriginalControlReturnProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
