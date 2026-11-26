"""ДЛЯ ТЕСТОВ. Сохраняет отрицательную проверку точной передачи момента до исправленного повторного расчёта."""
from pathlib import Path
from hashlib import sha256
import json,math

base=Path('D:/CodexScratch/RC_Section_NDM_TQ');old=base/'Cold1008PureBendingVolume_20261126';new=base/'Cold1010PureBendingVolumeExactInput_20261126'
prep=json.loads((old/'Preparation.json').read_bytes());result=json.loads((old/'Result.json').read_bytes())
expected=prep['section_target'];actual=result['section_target'];assert all(actual[i]==expected[i] for i in (0,2,3,4,5))
error=actual[1]-expected[1];assert error!=0 and abs(error)==math.ulp(expected[1])
correct=json.loads((new/'Result.json').read_bytes());assert correct['section_target']==expected and correct['termination'] is not None
proof=dict(exact_input_check_passed=False,excluded_from_exact_comparison=True,expected_section_target=expected,
    observed_section_target=actual,Mx_difference_Nmm=error,one_ulp=True,
    reason='PowerShell bare numeric argument rounded the Double while formatting the native process argument.',
    corrected_repeat='1010',corrected_repeat_target_bitexact=True,
    original_result_sha256=sha256((old/'Result.json').read_bytes()).hexdigest().upper(),
    physical_capacity_certified=False,whole_goal_completed=False)
(old/'InputMismatchProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ONE ULP INPUT MISMATCH PRESERVED; EXACT REPEAT1010 VERIFIED')
