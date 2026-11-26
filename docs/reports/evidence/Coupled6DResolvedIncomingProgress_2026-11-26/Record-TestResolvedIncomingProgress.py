"""ДЛЯ ТЕСТОВ. Архивирует958–960 с отрицательным уточнением959, направленными регрессиями и точными сохранёнными книгами."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json

here=Path(__file__).resolve().parent;repo=Path.cwd();base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold960ResolvedFieldProgress_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'WorkingGuideRegressions')
assert native['macros']==3 and native['failed']==0 and native['own_process_closed'] and not native['timeout600']
(root/'IncomingFocusedRegressionsProof.json').write_text(json.dumps(dict(native=native,
    scope='known envelope capacity and retained/discarded postpeak shoulder controls; not full ON/OFF/UI acceptance',
    full_acceptance=False,physical_capacity_certified=False),sort_keys=True,indent=2)+'\n',encoding='utf-8')
sources=[]
for registry,wanted in [('PathFollowingSources20261119.json','Abaqus2026ModifiedRiks.html'),
 ('AdaptiveStepSources20261120.json','OpenSees-DisplacementControl-20261120.html')]:
    folder=repo/'literature/Coupled6D/NumericalControl'
    source=next(item for item in json.loads((folder/registry).read_bytes()) if item['file']==wanted)
    assert sha256((folder/wanted).read_bytes()).hexdigest().upper()==source['sha256']
    sources.append(dict(source,registry=(folder/registry).relative_to(repo).as_posix(),primary_official_page_rechecked='2026-11-26'))
source=here/'Record-TestSequentialEventHistory.py';code=source.read_text('utf-8-sig')
old="roots={n:base/name for n,name in {955:'Cold955SequentialEventHistory_20261126',956:'Cold956RefinedSequentialHistory_20261126',957:'Cold957FourPartEventHistory_20261126'}.items()}"
new="roots={n:base/name for n,name in {958:'Cold958BoundedIncomingSteps_20261126',959:'Cold959AccurateIncomingSteps_20261126',960:'Cold960ResolvedFieldProgress_20261126'}.items()}"
changes={old:new,
 "keys={955:'SequentialEventHistoryProof.json',956:'RefinedSequentialHistoryProof.json',957:'FourPartEventHistoryProof.json'}":"keys={958:'BoundedIncomingStepsProof.json',959:'AccurateIncomingStepsProof.json',960:'ResolvedFieldProgressProof.json'}",
 "proof['original25_field_concrete_state_energy_stress_and_steel_stress_restored_bitexact_twice']":"proof['original25_field_concrete_state_energy_stress_and_steel_stress_restored_bitexact']",
 "out=repo/'docs/reports/evidence/Coupled6DSequentialEventHistory_2026-11-26'":"out=repo/'docs/reports/evidence/Coupled6DResolvedIncomingProgress_2026-11-26'",
 "publication_scope_verification=publication,":"publication_scope_verification=publication,primary_sources_rechecked=sources,focused_regressions=focused_regressions,",
 "print('ARCHIVED955-957'":"print('ARCHIVED958-960'",
}
for old,new in changes.items():assert code.count(old)==1;code=code.replace(old,new,1)
old="    files=fixed+[keys[n]]";new="    files=fixed+[keys[n],'IncomingFocusedRegressionsProof.json','Run-TestWorkingGuideRegressions.ps1','FieldProgressCriteriaBlock.txt']"
assert code.count(old)==1;code=code.replace(old,new,1)
old="'FourPartCompilationAncestor')";new="'FourPartCompilationAncestor','IncomingCompilationAncestor','AccurateIncomingCompilationAncestor','ResolvedFieldCompilationAncestor','PreRunEndpointSeedCorrection')"
assert code.count(old)==1;code=code.replace(old,new,1)
begin=code.index("helpers=set('''");end=code.index(";helpers.add(Path(__file__).name)",begin)
code=code[:begin]+"helpers=set('''Stage-TestBoundedIncomingSteps.py Check-TestBoundedIncomingSteps.py TestRunBoundedIncomingSteps.bas.txt\nStage-TestAccurateIncomingSteps.py Check-TestAccurateIncomingSteps.py Stage-TestResolvedFieldProgress.py\nCheck-TestResolvedFieldProgress.py TestRunFieldProgressSteps.bas.txt Rebuild-TestBookSegments.py\nRecord-TestSequentialEventHistory.py Finalize-TestExisting257Project.py Run-TestWorkingGuideRegressions.ps1\nSet-TestActiveStage.py Set-TestTerminalStage.py'''.split())"+code[end:]
code=code.replace('helpers.add(Path(__file__).name)',"helpers.add('Fix-TestIncomingRegressionDispatch.py');helpers.add(Path(__file__).name)",1)
begin=code.index('state.update(updated_local=');end=code.index('\ndata=(json.dumps(state,',begin)
code=code[:begin]+'''state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    latest_resolved_incoming_evidence=out.relative_to(repo).as_posix(),latest_bounded_incoming958=proofs['958'],
    latest_accurate_incoming959=proofs['959'],latest_resolved_field960=proofs['960'],latest_incoming_focused_regressions=focused_regressions,
    previous_goal_turn_classification='progress:958 one short step accepted with original guards;959 tighter internal solve retained1 step and exhausted budget, not promoted;960 tests resolved directed field continuation with original equilibrium/half/steel checks and stored lambda uncertainty, three negative guards and focused capacity/shoulder regressions. All outcomes preserved; physical capacity and cold gain unproven.')
'''+code[end:]
env={'__file__':str(Path(__file__).resolve()),'__name__':'resolved_incoming_archive_test','sources':sources,
    'focused_regressions':json.loads((root/'IncomingFocusedRegressionsProof.json').read_bytes())}
exec(compile(code,str(source),'exec'),env)
