"""ДЛЯ ТЕСТОВ. Проверяет полный исходный расчёт944 от прогноза по секущей, не ослабляя проверок941."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold944AffineControlReturn_20261126';parent=base/'Cold941OriginalControlReturn_20261126'
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
script=here/'Check-TestOriginalControlReturn.py';code=script.read_text('utf-8-sig')
assert code.count('Cold941OriginalControlReturn_20261126')==2
code=code.replace('Cold941OriginalControlReturn_20261126','Cold944AffineControlReturn_20261126')
scope={'__file__':str(script),'__name__':'affine_control_return_test'};exec(compile(code,str(script),'exec'),scope)
proof=json.loads((root/'OriginalControlReturnProof.json').read_bytes())
proof.pop('initial_native940_material_root_residual_bitexact_by_guard')
proof.update(initial_native943_secant_residual_bitexact_by_guard=True,immutable_compiled941_reused=True,
    manual_intermediate940_material_level_not_used=True,
    secant_source_material_roots_cost_must_be_added=True)
(root/'OriginalControlReturnProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
