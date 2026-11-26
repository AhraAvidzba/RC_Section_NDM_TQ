"""ДЛЯ ТЕСТОВ. Запрещает принимать выбор прогноза, теряющий уже подтверждённый переход исходного алгоритма на той же истории."""
from pathlib import Path
from hashlib import sha256
import json

base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
candidate=base/'Cold908ResidualArcPredictor_20261126';baseline=base/'Cold909ResidualArcPredictorBaseline_20261126'
a=json.loads((candidate/'ResidualArcPredictorProof.json').read_bytes())
b=json.loads((baseline/'ResidualArcPredictorBaselineProof.json').read_bytes())
assert a['source_states']==b['source_states']==32 and len(a['cases'])==len(b['cases'])==8
assert (candidate/'Input/Reverse1.bin').read_bytes()==(baseline/'Input/Reverse1.bin').read_bytes()
lost=[];gained=[]
for left,right in zip(a['cases'],b['cases']):
    for key in ('case','ratio','call_allowance'):assert left[key]==right[key]
    if right['solved'] and not left['solved']:lost.append(dict(case=left['case'],ratio=left['ratio'],allowance=left['call_allowance'],
        baseline_kind=right['selected_kind'],baseline_roots=right['independently_rebuilt_roots'],
        candidate_certificate=left['certificate'],candidate_predictor_comparison=left['predictor_comparison']))
    if left['solved'] and not right['solved']:gained.append(left['case'])
proof=dict(same_exact32_history=True,source_input_sha256=sha256((candidate/'Input/Reverse1.bin').read_bytes()).hexdigest().upper(),
 baseline_solved=sum(x['solved'] for x in b['cases']),candidate_solved=sum(x['solved'] for x in a['cases']),
 lost=lost,gained=gained,candidate_accepted=not lost,physical_capacity_certified=False,
 losing_valid_event_cannot_be_hidden_by164_synthetic_regressions=True)
(candidate/'ArcPredictorRegressionProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREDICTOR REGRESSION GATE',proof['baseline_solved'],'BASELINE',proof['candidate_solved'],'CANDIDATE;LOST',[x['case'] for x in lost],'GAINED',gained)
assert not lost,'Candidate rejected: a previously verified material-event transition was lost.'
