"""ДЛЯ ТЕСТОВ. Проверяет секущую945 от принятого состояния25 и одного материального равновесия с исходными защитами943."""
from pathlib import Path
from runpy import run_path
import json
here=Path(__file__).resolve().parent;script=here/'Check-TestAffineControlPredictor.py';code=script.read_text('utf-8-sig')
old="root=base/'Cold943AffineControlPredictor_20261126'";assert code.count(old)==1
code=code.replace(old,"root=base/'Cold945AcceptedAffinePredictor_20261126'",1)
old="low=v(root/'SourceLowMaterialRoot.bin')[2];high=v(root/'SourceHighMaterialRoot.bin')[2]\nratio=-low['values'][-1]/(high['values'][-1]-low['values'][-1]);assert 0<ratio<1\nintended=low['field']+ratio*(high['field']-low['field'])"
assert code.count(old)==1
code=code.replace(old,"high=v(root/'SourceHighMaterialRoot.bin')[2]\nleft=float(original[-1,:60]@accepted-meta[4]*meta[5])\nratio=-left/(high['values'][-1]-left);assert 0<ratio<1\nintended=accepted+ratio*(high['field']-accepted)",1)
scope={'__file__':str(script),'__name__':'accepted_affine_predictor_test'};exec(compile(code,str(script),'exec'),scope)
root=scope['root'];checked,fields=run_path(str(here/'Validate-TestHardPrepared.py'))['TestValidate'](root,'EnvelopeEnergy.Accepted')
assert len(fields)==25
proof=json.loads((root/'AffineControlPredictorProof.json').read_bytes())
proof.update(low_endpoint_is_already_accepted25=True,all25_accepted_equilibria_independently_reverified=True,
    low_auxiliary_material936_solve_not_used=True,only_one_additional_material_equilibrium_required=True)
(root/'AffineControlPredictorProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
