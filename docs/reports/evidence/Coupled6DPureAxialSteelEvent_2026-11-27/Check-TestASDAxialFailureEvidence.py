"""ДЛЯ ТЕСТОВ. Подтверждает отказ по стали при N=100 кН и отдельно сохраняет непройденную локализацию при N=10 кН."""
from pathlib import Path
from hashlib import sha256
import json, re

base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots = {1042: base / 'Cold1042ASDPureNBudget_20261127',
         1044: base / 'Cold1044ASDPureNColdRepeat_20261127'}
number = r'[-+]?\d+(?:\.\d+)?(?:[Ee][-+]?\d+)?'
events = {}
for case, root in roots.items():
    observed = json.loads((root / 'MainCapacityExecutionProof.json').read_bytes())
    external = json.loads((root / 'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
    diagnostic = (root / 'Native/MainCapacity.txt').read_text('utf-8-sig')
    assert observed['native']['own_process_closed'] and observed['has_steel_limit'] == 1 and observed['has_ultimate_point'] == 1
    assert external['all56plus4_equations_and_six_resultants_verified'] and external['reference_tolerance_violated_states'] == 0
    assert diagnostic.startswith('Status=FAIL;') and external['direction_N_Nmm'][0] == 100000.
    low, high = map(float, re.search(rf'Арматурное событие: lambda до = ({number}), lambda после = ({number})', diagnostic).groups())
    admissible, exceeded = map(float, re.search(rf'отношения деформаций к пределам: ({number}) и ({number})', diagnostic).groups())
    assert 0 < low <= high < 1 and admissible <= 1 < exceeded
    assert abs(low - observed['steel_limit_lambda']) < 1e-12
    assert abs(external['records'][-1]['steel_strain_utilization'] - admissible) < 1e-10
    events[str(case)] = dict(lambda_interval=[low, high], N_interval_N=[100000 * low, 100000 * high],
                            seconds=observed['elapsed_seconds'], accepted_states=external['accepted_states'],
                            admissible_utilization=admissible, exceeded_utilization=exceeded,
                            admissible_field_independently_replayed=True, exceeding_field_not_exported=True,
                            native_bracket_from_audited_solver=True,
                            history_sha256=sha256((root / 'Native/MainCapacityHistory.bin').read_bytes()).hexdigest().upper())
assert events['1042']['history_sha256'] == events['1044']['history_sha256']
reference_root = base / 'Cold1039ASDPureNTension_20261127'
driver = json.loads((reference_root / 'Native/IndependentAxialTRF.json').read_bytes())
checked = json.loads((reference_root / 'IndependentAxialTRFTestReplay/Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
assert driver['termination'] is None and checked['accepted_states'] == 43
assert checked['reference_tolerance_violated_states'] == 0
reference = driver['records'][-1]['force_N']
low, high = events['1042']['N_interval_N']
assert low <= reference <= high
negative = {}
for case, folder in [(1045, 'Cold1045ASDPureNScale10kN_20261127'), (1046, 'Cold1046ASDPureNEventRecovery_20261127')]:
    root = base / folder
    observed = json.loads((root / 'MainCapacityExecutionProof.json').read_bytes())
    external = json.loads((root / 'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
    assert observed['native']['own_process_closed'] and observed['has_steel_limit'] == 0
    assert observed['has_ultimate_point'] == 0 and external['reference_tolerance_violated_states'] == 0
    assert (root / 'Native/MainCapacity.txt').read_text('utf-8-sig').startswith('Status=NumFail;')
    negative[str(case)] = dict(maximum_N_N=10000 * external['maximum_lambda'],
                              last_steel_utilization=external['records'][-1]['steel_strain_utilization'],
                              accepted_states=external['accepted_states'], seconds=observed['elapsed_seconds'],
                              original_equilibrium_tolerances_satisfied=True,
                              steel_event_localization_passed=False, numerical_failure_not_physical_limit=True,
                              history_sha256=sha256((root / 'Native/MainCapacityHistory.bin').read_bytes()).hexdigest().upper())
assert negative['1045']['history_sha256'] == negative['1046']['history_sha256']
parent = Path('docs/reports/evidence/Coupled6DPureAxialFirstControls_2026-11-27')
pins = json.loads((parent / 'Files.json').read_bytes())
assert sha256((parent / 'RawEvidence.tar.xz').read_bytes()).hexdigest().upper() == pins['RawEvidence.tar.xz']
proof = dict(events=events, cold_repeat_history_bitexact=True, independent_N_at_steel_limit_N=reference,
             independent_reference_inside_native_bracket=True,
             lower_event_difference_from_independent_percent=100 * (low - reference) / reference,
             steel_failure_before_100kN_design_load_verified=True,
             independent_reference_archive=str(parent / 'RawEvidence.tar.xz'),
             independent_reference_archive_sha256=pins['RawEvidence.tar.xz'],
             negative_scaled_cases=negative, scale_event_acceptance_passed=False,
             global_concrete_capacity_unresolved=True, working_controller_not_promoted=True,
             full_acceptance_passed=False, whole_goal_completed=False)
(roots[1042] / 'AxialSteelFailureProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('100KN AXIAL STEEL FAILURE AND COLD REPEAT VERIFIED;10KN SCALE EVENT STILL FAILED;', proof['lower_event_difference_from_independent_percent'], 'PERCENT FROM INDEPENDENT REFERENCE')
