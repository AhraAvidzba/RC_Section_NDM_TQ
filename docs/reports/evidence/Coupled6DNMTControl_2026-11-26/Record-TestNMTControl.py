"""ДЛЯ ТЕСТОВ. Сохраняет пару NMT, исходные тензоры и точную книгу; отдельная проверка изгиба остаётся следующим этапом."""
from pathlib import Path
here=Path(__file__).resolve().parent
source=here/'Record-TestExternalComparisons.py';code=source.read_text('utf-8-sig')
begin=code.index('names={');end=code.index('\nroots=',begin)
code=code[:begin]+"names={1003:'Cold1003ASDNMTCold_20261126'}"+code[end:]
code=code.replace("keys={n:('FixedForceExternalProof.json' if n<1000 else 'ASDHardShearColdProof.json') for n in names}","keys={1003:'ASDHardShearColdProof.json'}",1)
anchor="out=repo/'docs/reports/evidence/Coupled6DExternalForceAndASD_2026-11-26'"
assert code.count(anchor)==1
code=code.replace(anchor,"""volume=base/'Cold1004OpenSeesNMTCore_20261126'
volume_proof=json.loads((volume/'NMTComparisonProof.json').read_bytes())
assert volume_proof['moment_gradient_from_Q_absent'] and not volume_proof['physical_capacity_certified']
for file in ('Preparation.json','Result.json','Run.log','Run.ndjson','Guard-IsolatedOffice.ps1','validate_opensees_hardshear.py','NMTComparisonProof.json'):
    raw['1004/'+file]=(volume/file).read_bytes()
out=repo/'docs/reports/evidence/Coupled6DNMTControl_2026-11-26'""",1)
begin=code.index("helpers=set('''");end=code.index(";todo=list(helpers)",begin)
code=code[:begin]+"helpers=set('''Record-TestNMTControl.py Record-TestExternalComparisons.py Stage-TestNMTControls.py Check-TestNMTComparison.py Rebuild-TestBookSegments.py Check-TestASDHardShearCold.py Replay-TestASDNativeMaterial.py Finalize-TestExisting257Project.py Set-TestActiveStage.py Set-TestTerminalStage.py'''.split())"+code[end:]
code=code.replace('all_seven_actual_books_rebuilt_bitexact=True','actual_NMT_test_book_rebuilt_bitexact=True,volume_proof=volume_proof',1)
code=code.replace('latest_external_force_and_ASD_evidence=','latest_nmt_control_evidence=',1).replace('latest_external_force_and_ASD_proofs=','latest_nmt_control_proofs=',1)
code=code.replace("print('ARCHIVED988-991/1000-1002'","print('ARCHIVED1003-1004'",1).replace('SEVEN EXACT BOOKS REBUILT','EXACT NMT BOOK REBUILT',1)
exec(compile(code,str(source),'exec'),{'__file__':str(Path(__file__).resolve()),'__name__':'nmt_control_archive_test'})
