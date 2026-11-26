"""ДЛЯ ТЕСТОВ. Сохраняет закрытый следующий этап N: арматурное событие, повтор, масштаб и попытку уточнения сжатия."""
from pathlib import Path

here = Path(__file__).resolve().parent
source = here / 'Record-TestASDPureAxialEvidence.py'
code = source.read_text('utf-8-sig')
start = code.index('roots = {')
end = code.index('\nout = ', start)
code = code[:start] + '''roots = {1042: base / 'Cold1042ASDPureNBudget_20261127',
         1043: base / 'Cold1043ASDPureNPeak_20261127',
         1044: base / 'Cold1044ASDPureNColdRepeat_20261127',
         1045: base / 'Cold1045ASDPureNScale10kN_20261127',
         1046: base / 'Cold1046ASDPureNEventRecovery_20261127'}''' + code[end:]
code = code.replace('Coupled6DPureAxialFirstControls_2026-11-27', 'Coupled6DPureAxialSteelEvent_2026-11-27')
code = code.replace("helpers = set('''", "helpers = set('''Record-TestASDAxialFollowup.py Stage-TestASDAxialController.py Check-TestASDAxialControllerSource.py Stage-TestASDAxialRepeat.py Check-TestASDAxialEvent.py Check-TestASDAxialFailureEvidence.py ", 1)
start = code.index("driver = json.loads(raw['1039/")
end = code.index('summary = dict(', start)
code = code[:start] + "event = json.loads(raw['1042/AxialSteelFailureProof.json'])\n" + code[end:]
start = code.index('               independent_axial_continuation=dict(')
end = code.index('               working_solver_not_changed=', start)
code = code[:start] + "               axial_steel_failure_proof=event,\n               scale_event_acceptance_passed=False,\n               next_stage='Pure N scale event and compression peak unresolved;1047 notincluded; then M/T/NMT/Q/NMTQ in user order',\n" + code[end:]
code = code.replace('all_three_actual_saved_books_included=True', 'all_five_actual_saved_books_included=True')
code = code.replace('PURE N FIRST CONTROLS SAVED', 'PURE N STEEL EVENT AND REPEATS SAVED')
code = code.replace('43 EXTERNAL STATES;NO GLOBAL PASS', 'ALL SAVED STATES REPLAYED;NO GLOBAL PASS')
exec(compile(code, str(source), 'exec'), dict(__file__=str(Path(__file__).resolve()), __name__='TestRecordAxialFollowup'))
