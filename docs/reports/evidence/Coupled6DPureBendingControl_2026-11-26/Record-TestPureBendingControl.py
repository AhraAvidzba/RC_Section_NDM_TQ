"""ДЛЯ ТЕСТОВ. Сохраняет контроль изгиба вместе с тайм-аутом, неудачной компиляцией и точным повтором входного момента."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Record-TestExternalComparisons.py';code=source.read_text('utf-8-sig')
begin=code.index('names={');end=code.index('\nroots=',begin)
code=code[:begin]+"""names={1005:'Cold1005PureBendingReference_20261126',1006:'Cold1006PureBendingCoupled_20261126',
1009:'Cold1009PureBendingASDOnly_20261126',1011:'Cold1011PureBendingASDCurvature_20261126',1012:'Cold1012PureBendingCDPCurvature_20261126',
1013:'Cold1013PureBendingPostCurveForce_20261126'}"""+code[end:]
begin=code.index('keys={');end=code.index('\nproofs=',begin)
code=code[:begin]+"""keys={1005:'PureBendingReferenceProof.json',1006:'PureBendingTimeoutPrefixProof.json',1009:'PureBendingCoupledProof.json',
1011:'PureBendingCurvatureProof.json',1012:'PureBendingCurvatureProof.json',1013:'PureBendingPostCurveForceProof.json'}"""+code[end:]
begin=code.index('for n in names:');end=code.index('\nparent_book=',begin)
code=code[:begin]+"""for n in names:
    proof=proofs[str(n)];assert proof['source_unchanged']
    if n==1005:assert proof['native']['own_process_closed'] and proof['planar_reference_physical_limit_verified']
    elif n==1006:assert proof['own_exact_excel_closed'] and not proof['whole_run_passed'] and proof['completed_equilibrium_verified']
    elif n==1009:assert proof['native']['own_process_closed'] and proof['all56plus4_equations_and_six_resultants_verified']
    elif n==1013:assert proof['native']['own_process_closed'] and proof['all_successful_forces_and56plus4_equations_independently_verified'] and not proof['physical_capacity_certified']
    else:assert proof['native']['own_process_closed'] and proof['all56plus4_equations_and_six_resultants_verified'] and not proof['physical_capacity_certified']
"""+code[end:]
begin=code.index("    files=set(prep['inputs'])");end=code.index('    for name in sorted(files):',begin)
code=code[:begin]+"""    files=set(prep['inputs'])|{'Preparation.json',keys[n],'docs/regression/CodeUpdate.json'}
    files.update(p.relative_to(root).as_posix() for p in root.iterdir() if p.is_file() and p.suffix in ('.json','.log','.ndjson','.ps1','.py','.txt'))
    for directory in ('Native','CompilationAttempt1Failure'):
        if (root/directory).exists():files.update(p.relative_to(root).as_posix() for p in (root/directory).rglob('*') if p.is_file())
"""+code[end:]
anchor="out=repo/'docs/reports/evidence/Coupled6DExternalForceAndASD_2026-11-26'";assert code.count(anchor)==1
code=code.replace(anchor,"""auxiliary_proofs={}
for n,name,key in ((1007,'Cold1007PureBendingFiberOpenSees_20261126','PureBendingFiberProof.json'),
                   (1008,'Cold1008PureBendingVolume_20261126','InputMismatchProof.json'),
                   (1010,'Cold1010PureBendingVolumeExactInput_20261126','PureBendingVolumeProof.json')):
    root=base/name;proof=json.loads((root/key).read_bytes());auxiliary_proofs[str(n)]=proof
    for path in root.rglob('*'):
        if path.is_file():raw[f'{n}/'+path.relative_to(root).as_posix()]=path.read_bytes()
replay=json.loads((roots[1011]/'Native/OpenSeesCurvatureMaterialReplay.json').read_bytes())
assert replay['material_tensor_match'] and not replay['section_equilibrium_solved_by_OpenSees']
out=repo/'docs/reports/evidence/Coupled6DPureBendingControl_2026-11-26'""",1)
begin=code.index("helpers=set('''");end=code.index(';todo=list(helpers)',begin)
code=code[:begin]+"""helpers=set('''Record-TestPureBendingControl.py Record-TestExternalComparisons.py Rebuild-TestBookSegments.py
Stage-TestPureBendingReference.py TestPureBendingPlanarReference.bas.txt Check-TestPureBendingReference.py
Stage-TestPureBendingCoupled.py TestPureBendingCoupled.bas.txt Check-TestPureBendingCoupled.py Check-TestPureBendingTimeoutPrefix.py
Stage-TestPureBendingASDOnly.py Stage-TestPureBendingCurvature.py TestPureBendingCurvature.bas.txt Check-TestPureBendingCurvature.py
Stage-TestPureBendingCDPCurvature.py Replay-TestASDCurvatureMaterial.py
Stage-TestPureBendingPostCurveForce.py Check-TestPureBendingPostCurveForce.py Check-TestPureBendingVolumeInputMismatch.py
Stage-TestPureBendingFiberOpenSees.py Run-TestPureBendingFiberOpenSees.py Check-TestPureBendingFiberOpenSees.py
Stage-TestPureBendingVolume.py Stage-TestPureBendingVolumeExactInput.py Check-TestPureBendingVolume.py
Finalize-TestExisting257Project.py Set-TestActiveStage.py Set-TestTerminalStage.py'''.split())"""+code[end:]
code=code.replace('all_seven_actual_books_rebuilt_bitexact=True','all_six_actual_books_rebuilt_bitexact=True,auxiliary_proofs=auxiliary_proofs,ASD_material_replay=replay',1)
# Полные тензоры остаются в сыром архиве; сводка хранит только компактную проверку.
code=code.replace("summary=dict(proofs=proofs", "replay={k:v for k,v in replay.items() if k!='observed_OpenSees_concrete_stresses'}\nsummary=dict(proofs=proofs",1)
code=code.replace('latest_external_force_and_ASD_evidence=','latest_pure_bending_control_evidence=',1).replace('latest_external_force_and_ASD_proofs=','latest_pure_bending_control_proofs=',1)
code=code.replace("print('ARCHIVED988-991/1000-1002'","print('ARCHIVED PURE BENDING1005-1013'",1).replace('SEVEN EXACT BOOKS REBUILT','SIX EXACT BOOKS REBUILT',1)
exec(compile(code,str(source),'exec'),{'__file__':str(Path(__file__).resolve()),'__name__':'pure_bending_archive_test'})
