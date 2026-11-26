"""ДЛЯ ТЕСТОВ. Подтверждает локализацию арматурного события N на двух масштабах и побайтную холодную повторяемость."""
from pathlib import Path
from hashlib import sha256
import json, re, struct

base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots = {1048: base / 'Cold1048ASDPureNEventSeed_20261127',
         1049: base / 'Cold1049ASDPureNSeedHighLoad_20261127',
         1050: base / 'Cold1050ASDPureNSeedColdRepeat_20261127'}
number = r'[-+]?\d+(?:\.\d+)?(?:[Ee][-+]?\d+)?'
events, book_hashes = {}, set()
for case, root in roots.items():
    observed = json.loads((root / 'MainCapacityExecutionProof.json').read_bytes())
    external = json.loads((root / 'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
    source = json.loads((root / 'EarlySourceProof.json').read_bytes())
    diagnostic = (root / 'Native/MainCapacity.txt').read_text('utf-8-sig')
    assert observed['native']['own_process_closed'] and observed['has_steel_limit'] == 1
    assert external['all56plus4_equations_and_six_resultants_verified'] and external['reference_tolerance_violated_states'] == 0
    assert source['controller_change_only'] == 'seed' and source['changed_components_exact']
    digest = sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()
    assert digest == source['saved_book_sha256']
    book_hashes.add(digest)
    packet = struct.unpack('<i6di', (root / 'Input/CapacityLoad.bin').read_bytes())
    force = packet[1]
    assert packet[0] == 6 and packet[-1] == 1 and all(value == 0 for value in packet[2:7])
    assert force == (100000. if case == 1049 else 10000.)
    assert external['direction_N_Nmm'] == [force, 0., 0., 0., 0., 0.]
    low, high = map(float, re.search(rf'Арматурное событие: lambda до = ({number}), lambda после = ({number})', diagnostic).groups())
    admissible, exceeded = map(float, re.search(rf'отношения деформаций к пределам: ({number}) и ({number})', diagnostic).groups())
    assert 0 < low < high and admissible < 1 < exceeded
    assert abs(low - observed['steel_limit_lambda']) < 1e-12
    assert abs(external['records'][-1]['steel_strain_utilization'] - admissible) < 1e-10
    if case == 1049:
        assert high < 1 and observed['has_ultimate_point'] == 1 and diagnostic.startswith('Status=FAIL;')
    else:
        assert low > 1 and observed['has_ultimate_point'] == 0 and diagnostic.startswith('Status=NumFail;')
        assert 'Определён арматурный предел. Для полного запаса требуется проверка предела бетонной ветви.' in diagnostic
        assert 'равновесие при lambda = 1 подтверждено: True' in diagnostic
    events[str(case)] = dict(base_N_N=force, lambda_interval=[low, high], N_interval_N=[force * low, force * high],
                            seconds=observed['elapsed_seconds'], accepted_states=external['accepted_states'],
                            admissible_utilization=admissible, exceeded_utilization=exceeded,
                            status=diagnostic.split(';')[0].split('=')[1],
                            admissible_field_independently_replayed=True, exceeding_field_not_exported=True,
                            history_sha256=sha256((root / 'Native/MainCapacityHistory.bin').read_bytes()).hexdigest().upper())
assert len(book_hashes) == 1
assert events['1048']['history_sha256'] == events['1050']['history_sha256']
reference_root = base / 'Cold1039ASDPureNTension_20261127'
driver = json.loads((reference_root / 'Native/IndependentAxialTRF.json').read_bytes())
reference = driver['records'][-1]['force_N']
assert driver['termination'] is None
for record in events.values():
    assert record['N_interval_N'][0] <= reference <= record['N_interval_N'][1]
left, right = events['1048']['N_interval_N'], events['1049']['N_interval_N']
assert max(left[0], right[0]) <= min(left[1], right[1])
failed = base / 'Cold1047ASDPureNEventCondition_20261127'
negative = json.loads((failed / 'MainCapacityExecutionProof.json').read_bytes())
assert negative['has_steel_limit'] == 0 and negative['native']['own_process_closed']
proof = dict(events=events, actual_saved_book_sha256=next(iter(book_hashes)),
             same_compiled_book_for_both_scales=True, cold_repeat_history_bitexact=True,
             independent_reference_N_N=reference, independent_reference_inside_all_native_brackets=True,
             two_scale_event_intervals_overlap=True, scale_event_localization_passed=True,
             steel_failure_before_100kN_design_load_verified=True,
             initial_state_for_10kN_verified=True, global_concrete_capacity_unresolved=True,
             production_controller_not_promoted=True, full_acceptance_passed=False, whole_goal_completed=False)
(roots[1048] / 'AxialSeedEventProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('AXIAL SEED EVENT VERIFIED;100KN FAILURE;10KN EVENT;COLD REPEAT BITEXACT;BRACKETS CONTAIN INDEPENDENT REFERENCE;NO GLOBAL PASS')
