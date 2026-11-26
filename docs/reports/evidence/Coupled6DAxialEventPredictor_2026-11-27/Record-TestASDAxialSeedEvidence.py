"""ДЛЯ ТЕСТОВ. Сохраняет исправленную локализацию N по решённым плечам вместе с отрицательным предыдущим контролем."""
from pathlib import Path

here = Path(__file__).resolve().parent
source = here / 'Record-TestASDPureAxialEvidence.py'
code = source.read_text('utf-8-sig')
start = code.index('roots = {')
end = code.index('\nout = ', start)
code = code[:start] + '''roots = {1047: base / 'Cold1047ASDPureNEventCondition_20261127',
         1048: base / 'Cold1048ASDPureNEventSeed_20261127',
         1049: base / 'Cold1049ASDPureNSeedHighLoad_20261127',
         1050: base / 'Cold1050ASDPureNSeedColdRepeat_20261127'}''' + code[end:]
code = code.replace('Coupled6DPureAxialFirstControls_2026-11-27', 'Coupled6DAxialEventPredictor_2026-11-27')
code = code.replace("helpers = set('''", "helpers = set('''Record-TestASDAxialSeedEvidence.py Stage-TestASDAxialController.py Check-TestASDAxialControllerSource.py Stage-TestASDAxialSeedRepeat.py Stage-TestASDAxialRepeat.py Check-TestASDAxialSeedEvidence.py Patch-TestAxialSteelSeed.py ", 1)
start = code.index("driver = json.loads(raw['1039/")
end = code.index('summary = dict(', start)
code = code[:start] + "event = json.loads(raw['1048/AxialSeedEventProof.json'])\nassert event['scale_event_localization_passed'] and event['cold_repeat_history_bitexact']\n" + code[end:]
start = code.index('               independent_axial_continuation=dict(')
end = code.index('               working_solver_not_changed=', start)
code = code[:start] + "               axial_seed_event_proof=event,\n               next_stage='Pure N compression peak remains unresolved; then M/T/NMT/Q/NMTQ in user order',\n" + code[end:]
code = code.replace('all_three_actual_saved_books_included=True', 'all_four_actual_saved_books_included=True')
code = code.replace('PURE N FIRST CONTROLS SAVED', 'AXIAL EVENT PREDICTOR AND REPEATS SAVED')
code = code.replace('43 EXTERNAL STATES;NO GLOBAL PASS', 'SCALE STEEL EVENT VERIFIED;NO GLOBAL PASS')
exec(compile(code, str(source), 'exec'), dict(__file__=str(Path(__file__).resolve()), __name__='TestRecordAxialSeedEvidence'))
