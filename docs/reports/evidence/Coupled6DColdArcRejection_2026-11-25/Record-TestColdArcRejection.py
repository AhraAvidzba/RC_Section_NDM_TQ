"""ДЛЯ ТЕСТОВ. Сохраняет завершённый холодный787, независимые56 состояний и точную причину отклонения дуги789–792."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={k:base/n for k,n in {'787':'Cold787ArcContextOrder_20261125','789':'Cold789RejectedOperationalArc_20261125',
    '790':'Cold790FixedContextColdReplay_20261125','791':'Cold791ResolvedLoadingEvent_20261125','792':'Cold792PreEventQuarter_20261125'}.items()}
proofs={k:json.loads((roots[k]/n).read_bytes()) for k,n in {'787':'OperationalColdProof.json','789':'RejectedOperationalArcProof.json',
    '790':'OperationalColdReplayProof.json','791':'RejectedOperationalArcProof.json','792':'PreEventQuarterProof.json'}.items()}
assert proofs['787']['cold']['capacity']=='NumFail' and proofs['787']['accepted_states']==56
assert proofs['789']['actual_rejection_due_to_unresolved_near_zero_event_confirmed']
assert proofs['790']['accepted_states']==56 and proofs['790']['source_cold_exact']
assert proofs['791']['event_resolved_at_inner_tolerance'] and not proofs['791']['returned_verified_candidate']
assert proofs['792']['quarter_solved'] and not proofs['792']['half_verified']
assert proofs['792']['field_half_error']<.005 and proofs['792']['lambda_half_error']>proofs['792']['lambda_half_bound']
native_check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
for key in ('789','790','791','792'):
    native,_=native_check(roots[key],'RealColdReplay');assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
native,_=native_check(roots['787'],'HBackendStock');assert native==dict(passed=28,failed=0,macros=1,timeout600=False,own_process_closed=True)
regression,_=native_check(roots['791'],'OperationalArcRegressions');assert regression==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
state_path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(state_path.read_bytes());assert not state['active_jobs']
out=repo/'docs/reports/evidence/Coupled6DColdArcRejection_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
for key,root in roots.items():
    for relative in ('OperationalColdProof.json','RejectedOperationalArcProof.json','OperationalColdReplayProof.json','PreEventQuarterProof.json',
        'Preparation.json','CompilePoint.log','CompilePointClosure.json','PointTraceSourceComposition.json','CompiledSourceReuse.json',
        'docs/regression/CodeUpdate.json','ParentCodeUpdate.json','ParentCompileClosure.json',
        'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledDeformationContinuation.cls',
        'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls','tests/modTestCoupledElasticState.bas',
        'Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1','Run-TestHMixedStock.ps1','Guard-IsolatedOffice.ps1',
        'tools/build_workbook/Update-IsolatedOptimizerCode.ps1'):
        p=root/relative
        if p.is_file():
            t=out/key/relative;t.parent.mkdir(parents=True,exist_ok=True);t.write_bytes(p.read_bytes())
    for directory in ('Input','Native'):
        for p in (root/directory).iterdir():
            if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx'):
                t=out/key/directory/p.name;t.parent.mkdir(parents=True,exist_ok=True);t.write_bytes(p.read_bytes())
helpers=('Stage-TestRejectedOperationalArc.py','Check-TestRejectedOperationalArc.py','Stage-TestFixedContextColdReplay.py',
    'Check-TestOperationalCold.py','Check-TestOperationalColdReplay.py','Stage-TestResolvedLoadingEvent.py',
    'Stage-TestPreEventQuarter.py','TestPreEventQuarter.bas.txt','Check-TestPreEventQuarter.py','Finalize-TestOperationalArc.py',
    'Run-TestOperationalArcRegressions.ps1','Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py','Compile-TestExternalPoint.ps1',
    'Finalize-TestStockYield.py','Audit-TestPrivateCoupledBook.py','Read-TestColdRawCapture.py','Check-TestActualReverseCapture.py',
    'Check-TestPersistentHEvidence.py','Validate-TestHardPrepared.py','Check-TestHardHalfSteps.py',Path(__file__).name)
for n in helpers:(out/n).write_bytes((here/n).read_bytes())
release={n:sha256((repo/n).read_bytes()).hexdigest().upper() for n in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
(out/'Summary.json').write_text(json.dumps(dict(proofs=proofs,regression791=regression,release_hashes=release,
    cold787_completed=True,all_started_stages_terminal=True,pre_event_candidate_not_admitted=True,
    physical_capacity_certified=False,whole_goal_completed=False),ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
    cold_arc_rejection_evidence='docs/reports/evidence/Coupled6DColdArcRejection_2026-11-25',
    cold787_final_summary={k:proofs['787'][k] for k in ('accepted_states','maximum_accepted_lambda','last_accepted_lambda','arc_calls','arc_recovered_directions')},
    previous_goal_turn_classification='progress: operational class/controller implemented and audited256; cold78560 and cold78756 equilibria independently verified, saved/reopened256 sources and164 regressions.789 exact cold10 rejection reproduced: event29/-1 nearly zero, half1.076e-11MPa below required9.728e-10; source already tighter than inner bound, so coarse source is not cause here.791 filters unresolved interpolation event before nearest choice, finds real15/+1, same physics and budgets;164 regressions pass.792 quarter solve gives field error0.0029745 but lambda error2.7788e-6 exceeds7.4525e-7, so no candidate admitted. All stages terminal, physical capacity unconfirmed.',
    next_action='Develop bounded repeated pre-event halving from same accepted history: reuse actual solved half field, solve fresh midpoint, check BOTH field0.005 and lambda errors plus force/gauges/source-yield/locality, keep common event30 primary and overall96 calls; reserve exact restoration.792 HALF NOT VERIFIED, do not integrate it as accepted. Expected need smaller candidate than half-event; test quarter/eighth/sixteenth with recorded probes and no threshold weakening. Integrate only verified fallback in CCoupledArcContinuation and cold controller, then new cold workflow, independent replay, step/mesh/external controls, all LoadPaths and full original deliverables.791 event filter is private; production sources/release unchanged. Full long acceptance remains NOT RUN.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (state_path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('787/789-792 ARCHIVED',len(pins),'PINS; ALL STARTED STAGES TERMINAL; PRE-EVENT STEP STILL REJECTED; GOAL ACTIVE')
