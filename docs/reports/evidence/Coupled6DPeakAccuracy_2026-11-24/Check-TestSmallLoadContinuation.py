"""ДЛЯ ТЕСТОВ. Проверяет исход713, включая грубый локальный пик, недостаточный для конечного статуса."""
from pathlib import Path
from runpy import run_path
import json,re


def TestCheck(root):
    """Сохраняет подтверждённые поля и ошибку повторения плеч; отсутствие физического предела остаётся явным."""
    here=Path(__file__).resolve().parent
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackTension')
    assert native==dict(passed=18,failed=1,macros=1,timeout600=False,own_process_closed=True)
    fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root,'PostCrackContinuation.Accepted')
    checkpoint=run_path(str(here/'Read-TestContinuationCheckpoint.py'))['TestRead'](root/'Native/Stock.Checkpoints.csv')
    lower=float(re.findall(r'наибольший допустимый lambda: ([0-9.E+-]+)',text)[0])
    assert abs(lower-1.44119402403768)<1e-13 and fields['states']==50
    assert 'проб: 58;' in text and 'проб при проверке плеч: 8' in text
    assert 'Суммарный спад найден, но его максимум не подтверждён повтором плеч.' in text
    law=root/'src/Materials/CCoupledConcreteResponse.cls'
    assert law.read_bytes()==(root.parent/'Cold711PriorDilationCurve_20261124/src/Materials/CCoupledConcreteResponse.cls').read_bytes()
    proof=dict(native=native,independent_fields=fields,last_complete_checkpoint=checkpoint,
        lower_lambda=lower,probes=58,shoulder_probes=8,early_local_peak_accuracy=.0260247217355122,
        original_capacity_accuracy=.001,coarse_local_peak_not_eligible_as_final_capacity=True,
        cumulative_decline_found=True,shoulder_verification_failed=True,
        load_change_resolution_formula='8*relativeForceTolerance*(1+abs(lambda)); not capacityAccuracy',
        original_material711_exact=True,physical_capacity_certified=False,entire_goal_completed=False)
    (root/'SmallLoadContinuationProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('713 VERIFIED50FIELDS;58PROBES;8SHOULDERS; NO FINAL CAPACITY; LAST CHECKPOINT',checkpoint['sequence'])
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold713SmallLoadContinuation_20261124'))
