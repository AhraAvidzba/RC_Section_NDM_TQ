"""ДЛЯ ТЕСТОВ. Сохраняет подтверждённое ограничение стали ASD и незавершённый получасовой CDP-контроль без общего PASS."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Record-TestExternalComparisons.py';code=source.read_text('utf-8-sig')
begin=code.index('names={');end=code.index('\nroots=',begin)
code=code[:begin]+"""names={1014:'Cold1014PureBendingCDPLong_20261126',1015:'Cold1015PureBendingSteelLimit_20261126',
1016:'Cold1016PureBendingSteelLimitFixed_20261126'}"""+code[end:]
begin=code.index('keys={');end=code.index('\nproofs=',begin)
code=code[:begin]+"keys={1014:'PureBendingCurvatureProof.json',1015:'SteelLimitDriverFailureProof.json',1016:'PureBendingSteelLimitProof.json'}"+code[end:]
code=code.replace("assert not state['active_jobs'] and state['active_exec_session'] is None and state['goal_status']=='active'",
    "assert state['goal_status']=='active' and all(Path(j['root']).name not in names.values() for j in state['active_jobs'])",1)
begin=code.index('for n in names:');end=code.index('\nparent_book=',begin)
code=code[:begin]+"""for n in names:
    proof=proofs[str(n)];assert proof['source_unchanged']
    if n==1014:assert proof['native']['own_process_closed'] and proof['native']['timeout1800'] and not proof['whole_run_passed']
    elif n==1015:assert proof['own_exact_excel_closed'] and proof['test_driver_failed_before_steel_control'] and not proof['whole_run_passed']
    else:assert proof['native']['own_process_closed'] and proof['longitudinal_strain_limit_equilibrium_confirmed'] and proof['global_peak_not_certified']
replay=json.loads((roots[1016]/'Native/OpenSeesSteelLimitMaterialReplay.json').read_bytes())
assert replay['material_tensor_match'] and not replay['section_equilibrium_solved_by_OpenSees']
"""+code[end:]
begin=code.index("    files=set(prep['inputs'])");end=code.index('    for name in sorted(files):',begin)
code=code[:begin]+"""    files=set(prep['inputs'])|{'Preparation.json',keys[n],'docs/regression/CodeUpdate.json'}
    files.update(p.relative_to(root).as_posix() for p in root.iterdir() if p.is_file() and p.suffix in ('.json','.log','.ndjson','.ps1','.py','.txt'))
    files.update(p.relative_to(root).as_posix() for p in (root/'Native').rglob('*') if p.is_file())
"""+code[end:]
code=code.replace("out=repo/'docs/reports/evidence/Coupled6DExternalForceAndASD_2026-11-26'",
    "out=repo/'docs/reports/evidence/Coupled6DBendingMaterialLimits_2026-11-26'",1)
begin=code.index("helpers=set('''");end=code.index(';todo=list(helpers)',begin)
code=code[:begin]+"""helpers=set('''Record-TestBendingMaterialLimits.py Record-TestExternalComparisons.py Rebuild-TestBookSegments.py
Stage-TestPureBendingCDPLong.py Check-TestPureBendingCDPLong.py Check-TestPureBendingCurvature.py
Stage-TestPureBendingSteelLimit.py Stage-TestPureBendingSteelLimitFixed.py Check-TestPureBendingSteelLimit.py
Stage-TestPureBendingPostCurveForce.py Replay-TestASDCurvatureMaterial.py Set-TestActiveStage.py Set-TestTerminalStage.py
Finalize-TestExisting257Project.py'''.split())"""+code[end:]
code=code.replace('all_seven_actual_books_rebuilt_bitexact=True','all_three_actual_books_rebuilt_bitexact=True,ASD_material_replay={k:v for k,v in replay.items() if k!="observed_OpenSees_concrete_stresses"}',1)
code=code.replace('all_started_stages_terminal=True','archived_started_stages_terminal=True,other_jobs_still_active=state["active_jobs"]',1)
code=code.replace('latest_external_force_and_ASD_evidence=','latest_bending_material_limits_evidence=',1).replace('latest_external_force_and_ASD_proofs=','latest_bending_material_limits_proofs=',1)
code=code.replace("print('ARCHIVED988-991/1000-1002'","print('ARCHIVED1014-1016 BENDING MATERIAL LIMITS'",1).replace('SEVEN EXACT BOOKS REBUILT','THREE EXACT BOOKS REBUILT',1)
exec(compile(code,str(source),'exec'),{'__file__':str(Path(__file__).resolve()),'__name__':'bending_material_limit_archive_test'})
