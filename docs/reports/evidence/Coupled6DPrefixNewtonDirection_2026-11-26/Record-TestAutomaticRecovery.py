"""ДЛЯ ТЕСТОВ. Архивирует направленные проверки без будущего корня, включая отрицательные варианты и точные сохранённые книги."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser

here=Path(__file__).resolve().parent
p=ArgumentParser(description=__doc__);p.add_argument('--group',choices=['arc','prefix','capture'],required=True);a=p.parse_args()
if a.group=='arc':
    names={961:'Cold961FailedPredictorEvent_20261126',962:'Cold962CurrentOperationalArc_20261126',
        963:'Cold963IncomingEventBudget_20261126',964:'Cold964RequestedArcAccuracy_20261126'}
    keys={961:'FailedPredictorEventProof.json',962:'CurrentOperationalArcProof.json',963:'IncomingEventBudgetProof.json',964:'RequestedArcAccuracyProof.json'}
    suffix='Coupled6DAutomaticEventRecovery_2026-11-26'
elif a.group=='prefix':
    names={965:'Cold965ControlledSecantPrefix_20261126',966:'Cold966ShortSecantPrefix_20261126',967:'Cold967LongerSecantPrefix_20261126'}
    keys={n:'ControlledSecantPrefixProof.json' for n in names};suffix='Coupled6DOriginalScalarPrefixes_2026-11-26'
else:
    names={968:'Cold968PrefixNewtonCapture_20261126',969:'Cold969PrefixSpatialFloor_20261126',970:'Cold970PrefixSurfaceMatrices_20261126'}
    keys={968:'PrefixNewtonCaptureProof.json',969:'PrefixNewtonCaptureProof.json',970:'SurfaceMatricesProof.json'}
    suffix='Coupled6DPrefixNewtonDirection_2026-11-26'
source=here/'Record-TestSequentialEventHistory.py';code=source.read_text('utf-8-sig')
changes={
    "roots={n:base/name for n,name in {955:'Cold955SequentialEventHistory_20261126',956:'Cold956RefinedSequentialHistory_20261126',957:'Cold957FourPartEventHistory_20261126'}.items()}":
        'roots={n:base/name for n,name in '+repr(names)+'.items()}',
    "keys={955:'SequentialEventHistoryProof.json',956:'RefinedSequentialHistoryProof.json',957:'FourPartEventHistoryProof.json'}":'keys='+repr(keys),
    "    assert proof['original25_field_concrete_state_energy_stress_and_steel_stress_restored_bitexact_twice']":
        "    assert not proof.get('cold_gain_claimed',False) # Проверка продолжения не доказывает предел или улучшение полного холодного расчёта.",
    "out=repo/'docs/reports/evidence/Coupled6DSequentialEventHistory_2026-11-26'":
        "out=repo/'docs/reports/evidence/"+suffix+"'",
    "'FourPartCompilationAncestor')":
        "'FourPartCompilationAncestor','IncomingEventBudgetCompilationAncestor','RequestedArcAccuracyCompilationAncestor','SecantPrefixCompilationAncestor','LongerSecantCompilationAncestor','PrefixSpatialFloorCompilationAncestor')",
    "    files=fixed+[keys[n]]":
        "    files=fixed+[keys[n]]+[p.name for p in root.iterdir() if p.is_file() and p.suffix in ('.json','.py','.txt','.ps1','.bin') and p.name not in ('Compile.log',)]",
    'all3_saved_books_rebuilt_bitexact_from_archived_segments=True':'all_saved_books_rebuilt_bitexact_from_archived_segments=True',
    "print('ARCHIVED955-957'":"print('ARCHIVED_AUTOMATIC_RECOVERY'",
}
for old,new in changes.items():assert code.count(old)==1,old;code=code.replace(old,new,1)
begin=code.index("helpers=set('''");end=code.index(";helpers.add(Path(__file__).name)",begin)
code=code[:begin]+"""helpers=set('''Stage-TestFailedPredictorEvent.py Check-TestFailedPredictorEvent.py Stage-TestCurrentOperationalArc.py
Check-TestCurrentOperationalArc.py Stage-TestIncomingEventBudget.py Check-TestIncomingEventBudgetScope.py
Stage-TestRequestedArcAccuracy.py Check-TestRequestedArcAccuracyScope.py Stage-TestControlledSecantPrefix.py
Check-TestControlledSecantPrefix.py TestTryControlledSecantPrefix.cls.txt TestRunControlledSecantPrefix.bas.txt
Stage-TestShortSecantPrefix.py Stage-TestLongerSecantPrefix.py Rebuild-TestBookSegments.py
Stage-TestPrefixNewtonCapture.py Check-TestPrefixNewtonCapture.py Stage-TestPrefixSpatialFloor.py
Stage-TestPrefixSurfaceMatrices.py Check-TestPrefixSurfaceMatrices.py Check-TestSurfaceMatrices.py
Record-TestSequentialEventHistory.py Finalize-TestExisting257Project.py Set-TestActiveStage.py Set-TestTerminalStage.py'''.split())"""+code[end:]
begin=code.index('state.update(updated_local=');end=code.index('\ndata=(json.dumps(state,',begin)
code=code[:begin]+'''state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    latest_automatic_recovery_evidence_group=out.relative_to(repo).as_posix())
state.setdefault('automatic_recovery_evidence_groups',{})[GROUP]=out.relative_to(repo).as_posix()
state.setdefault('latest_automatic_recovery_proofs',{}).update(proofs)
'''+code[end:]
code=code.replace("print('ARCHIVED_AUTOMATIC_RECOVERY'","print('ARCHIVED_AUTOMATIC_RECOVERY',GROUP",1)
code=code.replace('ALL3 ACTUAL BOOKS REBUILT BITEXACT','ALL ACTUAL BOOKS REBUILT BITEXACT',1)
# Итоговый архив обязан оставаться ниже жёсткого лимита отдельного GitHub-файла.
anchor="checked={};actual_pins=None;actual_parts={}"
code=code.replace(anchor,"assert archivefile.stat().st_size<100*1024**2, 'Разбить архив перед публикацией: больше100MiB.'\n"+anchor,1)
env={'__file__':str(Path(__file__).resolve()),'__name__':'automatic_recovery_archive_test','GROUP':a.group}
exec(compile(code,str(source),'exec'),env)
