"""ДЛЯ ТЕСТОВ. Проверяет максимум сжатия, холодный повтор, масштаб нагрузки и независимые ветви при исходных допусках."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json, struct

base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
names = {1051: 'Cold1051ASDPureNCurrentCompression_20261127',
         1052: 'Cold1052ASDPureNPeakObservation_20261127',
         1053: 'Cold1053ASDPureNPeakSamples_20261127',
         1054: 'Cold1054ASDIndependentCompressionPeak_20261127',
         1055: 'Cold1055ASDPureNAdaptivePeak_20261127',
         1056: 'Cold1056ASDPureNFullCompressionBranch_20261127',
         1057: 'Cold1057ASDIndependentCompressionFine_20261127',
         1058: 'Cold1058ASDPureNCompressionOverload_20261127'}
roots = {case: base / name for case, name in names.items()}
native_cases = (1051, 1052, 1053, 1055, 1056, 1058)


def TestJSON(path):
    """Читает точный сохранённый результат; отсутствие доказательства не подменяет успешным значением."""
    return json.loads(path.read_bytes())


def TestDigest(path):
    """Вычисляет контрольную сумму фактических байтов без нормализации."""
    return sha256(path.read_bytes()).hexdigest().upper()


def TestReplay(root):
    """Проверяет полноту независимого повтора по всем уравнениям, шести усилиям и реальным исходным допускам."""
    result = TestJSON(root / 'Native/OpenSeesMainCapacityEquilibrium.json')
    assert result['all56plus4_equations_and_six_resultants_verified']
    assert result['reference_tolerance_violated_states'] == 0
    assert result['input_history_sha256'] == TestDigest(root / 'Native/MainCapacityHistory.bin')
    assert result['declared_absolute_tolerance_MPa'] == result['declared_relative_tolerance'] == 5e-7
    assert result['full_steel_plastic_history_replayed'] and result['actual_OpenSees_material_replay']
    assert result['section_equilibrium_not_solved_by_OpenSees']
    return result


observed = {}
for case in native_cases:
    root = roots[case]
    result = TestJSON(root / 'MainCapacityExecutionProof.json')
    assert result['native']['own_process_closed'] and result['source_unchanged']
    source = TestJSON(root / 'EarlySourceProof.json')
    assert source['changed_components_exact'] and source['full_equilibrium_acceptance_unchanged']
    assert source['saved_book_sha256'] == TestDigest(root / 'docs/regression/Diagnostic.xlsm')
    TestReplay(root)
    observed[case] = result
for case in roots:
    assert TestDigest(roots[case] / 'Input/MaterialContract.json') == TestDigest(roots[1051] / 'Input/MaterialContract.json')
    assert TestDigest(roots[case] / 'Native/MainCapacity.Operators.csv') == TestDigest(roots[1051] / 'Native/MainCapacity.Operators.csv')
assert len({TestDigest(roots[case] / 'Native/MainCapacityHistory.bin') for case in (1051, 1052, 1053)}) == 1
assert TestDigest(roots[1055] / 'Native/MainCapacityHistory.bin') == TestDigest(roots[1056] / 'Native/MainCapacityHistory.bin')
assert TestDigest(roots[1056] / 'docs/regression/Diagnostic.xlsm') == TestDigest(roots[1058] / 'docs/regression/Diagnostic.xlsm')
for case in (1051, 1052, 1053):
    assert observed[case]['has_ultimate_point'] == 0 and 'Status=NumFail;' in observed[case]['diagnostics']
    assert observed[case]['peak_accuracy'] == observed[case]['peak_lambda'] == 0
for case in (1052, 1053):
    rows = (roots[case] / 'Native/PeakRefinement.Values.csv').read_text('utf-8-sig').splitlines()
    completion, = [row.split(';') for row in rows if row.startswith('completion;')]
    assert float(completion[3]) > float(completion[4]) == .001

values = {}
for case, expected_load in ((1055, -100000.), (1056, -100000.), (1058, -1000000.)):
    root = roots[case]
    packet = struct.unpack('<i6di', (root / 'Input/CapacityLoad.bin').read_bytes())
    assert packet == (6, expected_load, 0., 0., 0., 0., 0., 1)
    result = observed[case]
    assert result['has_ultimate_point'] == 1 and result['peak_accuracy'] <= .001
    force = abs(expected_load) * result['lambda_capacity']
    uncertainty = abs(expected_load) * result['peak_accuracy']
    assert result['peak_lambda'] == result['lambda_capacity']
    values[case] = dict(capacity_N=force, estimated_error_N=uncertainty,
                        interval_N=[force - uncertainty, force + uncertainty],
                        elapsed_seconds=result['elapsed_seconds'], lambda_capacity=result['lambda_capacity'])
assert 'Status=OK;' in observed[1056]['diagnostics']
assert 'Status=FAIL;' in observed[1058]['diagnostics']
assert observed[1058]['lambda_capacity'] + observed[1058]['peak_accuracy'] < 1
assert max(values[c]['interval_N'][0] for c in values) <= min(values[c]['interval_N'][1] for c in values)

tails = {}
reader_class = run_path(str(Path(__file__).resolve().parent / 'Read-TestColdRawCapture.py'))['TestReader']
for case, count in ((1056, 56), (1058, 26)):
    root = roots[case] / 'FullCompressionBranchTestReplay'
    replay = TestReplay(root)
    assert replay['accepted_states'] == count
    records = replay['records']
    assert records[-1]['lambda_value'] < replay['maximum_lambda']
    assert .999 < records[-1]['steel_strain_utilization'] < 1
    assert all(record['steel_strain_utilization'] <= 1 for record in records)
    assert abs(records[-1]['concrete_min_longitudinal_strain']) > .0179
    raw = (root / 'Native/MainCapacityHistory.bin').read_bytes()
    reader = reader_class(raw)
    for i in range(4):
        reader.VectorTest()
    strains = []
    for i in range(count):
        reader.VectorTest()
        strains.append(abs(reader.VectorTest()[0]))
    assert all(right >= left for left, right in zip(strains, strains[1:]))
    assert reader.position == len(raw)
    tails[case] = dict(accepted_states=count, monotonic_axial_strain=True,
                       final_concrete_min_strain=records[-1]['concrete_min_longitudinal_strain'],
                       final_steel_utilization=records[-1]['steel_strain_utilization'],
                       maximum_force_error_N=replay['maximum_force_error_N'],
                       all_equilibrium_checks_passed=True)

external = {}
for case, replay_name, expected_states in ((1054, 'IndependentCompressionPeakTestReplay', 58),
                                           (1057, 'IndependentCompressionFineTestReplay', 62)):
    root = roots[case]
    result = TestJSON(root / 'Native/IndependentCompressionPeakTRF.json')
    replay = TestReplay(root / replay_name)
    assert replay['accepted_states'] == expected_states and result['termination'] is None
    assert all(record['accepted'] for record in result['records'])
    peak = max(result['records'], key=lambda record: record['lambda_value'])
    assert result['records'][-1]['lambda_value'] < peak['lambda_value']
    assert result['independent_OpenSees_structural_analysis'] is False
    force = abs(peak['force_N'])
    external[case] = dict(capacity_N=force, axial_strain=peak['axial_strain'],
                          elapsed_seconds=result['elapsed_seconds'], inherited_states=result['inherited_states'],
                          new_states=len(result['records']), replayed_states=expected_states,
                          all_equilibrium_checks_passed=True)
fine = external[1057]['capacity_N']
assert values[1056]['interval_N'][0] <= fine <= values[1056]['interval_N'][1]
proof = dict(scope='Pure compression ASD; chosen materials, discretization and followed strain-controlled branch',
             cold_repeat_history_bitexact=True, observation_did_not_change_history=True,
             original_equilibrium_and_lambda_tolerances_not_relaxed=True,
             peak_load_capacity_confirmed_on_followed_branch=True,
             overloaded_case_physical_FAIL_confirmed=True, scale_intervals_overlap=True,
             native_cases=values, full_tails=tails, independent_curves=external,
             fine_minus_native_N=fine - values[1056]['capacity_N'],
             fine_minus_native_percent=100 * (fine / values[1056]['capacity_N'] - 1),
             public_solver_not_updated=True, pure_N_stage_complete=False,
             CDP_and_planar_N_and_direct_state_checks_pending=True,
             full_acceptance_not_run=True, full_run_10second_target_not_proved=True,
             independent_OpenSees_structural_analysis=False, whole_goal_completed=False)
(roots[1056] / 'AxialCompressionLimitProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('AXIAL COMPRESSION LIMIT VERIFIED', json.dumps(proof, sort_keys=True))
