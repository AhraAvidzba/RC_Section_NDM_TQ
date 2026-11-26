"""ДЛЯ ТЕСТОВ. Проверяет уменьшенный шаг948 без выдачи его за решение полной управляющей цели947."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;script=here/'Check-TestWorkingSecantReturn.py';code=script.read_text('utf-8-sig')
assert code.count('Cold947WorkingSecantReturn_20261126')==2
code=code.replace('Cold947WorkingSecantReturn_20261126','Cold948QuarterWorkingReturn_20261126')
scope={'__file__':str(script),'__name__':'quarter_working_return_test'};exec(compile(code,str(script),'exec'),scope)
root=scope['root'];proof=json.loads((root/'WorkingSecantReturnProof.json').read_bytes())
proof.update(immutable_compiled947_reused=True,reduced_control_increment_fraction=.25,
    full_original_target_not_claimed_solved=True,all_original_tolerances_and_guards_preserved=True)
(root/'WorkingSecantReturnProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
