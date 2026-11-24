"""ДЛЯ ТЕСТОВ. Проверяет фактический автоматический поиск712 и все принятые поля без назначения неустановленного предела."""
from pathlib import Path
from runpy import run_path
import json,re


def TestCheck(root):
    """Лимит проб сохраняется как численный исход; независимая сборка подтверждает только принятые состояния."""
    here=Path(__file__).resolve().parent
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackTension')
    assert native==dict(passed=18,failed=1,macros=1,timeout600=False,own_process_closed=True)
    fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root,'PostCrackContinuation.Accepted')
    lower=float(re.findall(r'наибольший допустимый lambda: ([0-9.E+-]+)',text)[0])
    assert abs(lower-1.44119480550414)<1e-13 and 'проб: 96;' in text
    assert (root/'src/Materials/CCoupledConcreteResponse.cls').read_bytes()==(root.parent/'Cold711PriorDilationCurve_20261124/src/Materials/CCoupledConcreteResponse.cls').read_bytes()
    assert (root/'src/Solver/CCoupledDeformationContinuation.cls').read_bytes()==(root.parent/'Cold697PostCrackRankTrust_20261124/src/Solver/CCoupledDeformationContinuation.cls').read_bytes()
    proof=dict(native=native,independent_fields=fields,confirmed_lower_lambda=lower,probes=96,
       shoulder_probes=39,known_uniaxial_peak_lambda=172967.56279474715/120000,
       physical_capacity_certified=False,original_probe_and_accuracy_limits_preserved=True,
       current_material711_exact=True,adaptive_controller697_exact=True,entire_goal_completed=False)
    (root/'PriorDilationCapacityProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('712 VERIFIED ACCEPTED STATES',fields['states'],'LOWER',lower,'LAST',fields['point_fields'][-1],'CAPACITY NOT CONFIRMED')
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold712PriorDilationCapacity_20261124'))
