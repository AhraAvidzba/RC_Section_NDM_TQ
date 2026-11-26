"""ДЛЯ ТЕСТОВ. Сохраняет завершённый по времени CDP-контроль5400с и независимое равновесие полного записанного префикса."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Record-TestExternalComparisons.py';code=source.read_text('utf-8-sig')
begin=code.index('names={');end=code.index('\nroots=',begin)
code=code[:begin]+"names={1019:'Cold1019PureBendingCDP90_20261126'}"+code[end:]
begin=code.index('keys={');end=code.index('\nproofs=',begin)
code=code[:begin]+"keys={1019:'PureBendingCurvatureProof.json'}"+code[end:]
code=code.replace("assert not state['active_jobs'] and state['active_exec_session'] is None and state['goal_status']=='active'",
    "assert state['goal_status']=='active' and all(Path(j['root']).name not in names.values() for j in state['active_jobs'])",1)
begin=code.index('for n in names:');end=code.index('\nparent_book=',begin)
code=code[:begin]+'''for n in names:
    proof=proofs[str(n)];assert proof['source_unchanged'] and proof['native']['own_process_closed']
    assert proof['native']['timeout5400'] and proof['completed_prefix_only'] and not proof['whole_run_passed']
    assert proof['all56plus4_equations_and_six_resultants_verified'] and not proof['physical_capacity_certified']
'''+code[end:]
begin=code.index("    files=set(prep['inputs'])");end=code.index('    for name in sorted(files):',begin)
code=code[:begin]+'''    files=set(prep['inputs'])|{'Preparation.json',keys[n],'docs/regression/CodeUpdate.json'}
    files.update(p.relative_to(root).as_posix() for p in root.iterdir() if p.is_file() and p.suffix in ('.json','.log','.ndjson','.ps1','.py','.txt'))
    files.update(p.relative_to(root).as_posix() for p in (root/'Native').rglob('*') if p.is_file())
'''+code[end:]
code=code.replace("out=repo/'docs/reports/evidence/Coupled6DExternalForceAndASD_2026-11-26'",
    "out=repo/'docs/reports/evidence/Coupled6DCDP90_2026-11-26'",1)
begin=code.index("helpers=set('''");end=code.index(';todo=list(helpers)',begin)
code=code[:begin]+"""helpers=set('''Record-TestCDP90.py Record-TestExternalComparisons.py Rebuild-TestBookSegments.py
Stage-TestPureBendingCDP90.py Check-TestPureBendingCDP90.py Check-TestPureBendingCDPLong.py Check-TestPureBendingCurvature.py
Stage-TestPureBendingCDPLong.py Set-TestActiveStage.py Set-TestTerminalStage.py'''.split())"""+code[end:]
code=code.replace('all_seven_actual_books_rebuilt_bitexact=True','actual_CDP90_book_rebuilt_bitexact=True',1)
code=code.replace('all_started_stages_terminal=True','archived_started_stage_terminal=True,other_jobs_still_active=state["active_jobs"]',1)
code=code.replace('latest_external_force_and_ASD_evidence=','latest_CDP90_evidence=',1).replace('latest_external_force_and_ASD_proofs=','latest_CDP90_proofs=',1)
code=code.replace("print('ARCHIVED988-991/1000-1002'","print('ARCHIVED1019 CDP90 TIMEOUT WITH VERIFIED PREFIX'",1).replace('SEVEN EXACT BOOKS REBUILT','EXACT CDP90 BOOK REBUILT',1)
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='CDP90_archive_test'))
