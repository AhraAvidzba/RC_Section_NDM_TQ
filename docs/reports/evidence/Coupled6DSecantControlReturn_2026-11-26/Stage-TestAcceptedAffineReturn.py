"""ДЛЯ ТЕСТОВ. Переиспользует проверенный путь944 для секущей от принятого состояния25, без нижнего материального расчёта."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;script=here/'Stage-TestAffineControlReturn.py';code=script.read_text('utf-8-sig')
for before,after in [('Cold943AffineControlPredictor_20261126','Cold945AcceptedAffinePredictor_20261126'),
    ('Cold944AffineControlReturn_20261126','Cold946AcceptedAffineReturn_20261126')]:
    assert code.count(before)==1;code=code.replace(before,after,1)
scope={'__file__':str(script),'__name__':'accepted_affine_return_stage_test'};exec(compile(code,str(script),'exec'),scope)
root=scope['root'];prep=json.loads((root/'Preparation.json').read_bytes())
prep['directed_scope'].update(case='Original61ReturnFromAccepted25Secant',low_material936_solve_not_used=True,
    low_endpoint_is_accepted25=True,material938_high_root_steps=17,full_material_plus_return_effort_required=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
