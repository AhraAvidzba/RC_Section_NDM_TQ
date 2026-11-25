"""ДЛЯ ТЕСТОВ. Архивирует проверенный подход до события и отрицательные холодные результаты793–801."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd(); here=Path(__file__).resolve().parent
base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={793:'Cold793AdaptivePreEvent_20261125',794:'Cold794PreEventBudget_20261125',
 795:'Cold795PreEventSixCalls_20261125',796:'Cold796PreEventThreeCalls_20261125',
 797:'Cold797ArcGrowthControl_20261125',798:'Cold798AdaptiveColdReplay_20261125',
 799:'Cold799ArcGrowthColdReplay_20261125',800:'Cold800PostArcAccuracy_20261125',
 801:'Cold801PostArcAccuracyColdReplay_20261125'}
roots={k:base/v for k,v in names.items()}
selected={793:['AdaptivePreEventProof.json'],794:['AdaptivePreEventProof.json','OperationalColdProof.json'],
 795:['AdaptivePreEventProof.json'],796:['AdaptivePreEventProof.json'],797:['OperationalColdProof.json'],
 798:['OperationalColdReplayProof.json'],799:['OperationalColdReplayProof.json'],
 800:['OperationalColdProof.json'],801:['OperationalColdReplayProof.json']}
proofs={str(k):{n:json.loads((roots[k]/n).read_bytes()) for n in ns} for k,ns in selected.items()}
for k in (793,794):
 assert proofs[str(k)]['AdaptivePreEventProof.json']['returned_verified_candidate']
for k in (795,796):
 assert not proofs[str(k)]['AdaptivePreEventProof.json']['returned_verified_candidate']
for k,count in ((794,67),(797,47),(800,45)):
 p=proofs[str(k)]['OperationalColdProof.json']
 assert p['accepted_states']==count and p['arc_recovered_directions']==1
 assert p['cold']['capacity']=='NumFail' and p['cold']['limit']==''
for k,count in ((798,67),(799,47),(801,45)):
 p=proofs[str(k)]['OperationalColdReplayProof.json']
 assert p['accepted_states']==count and p['source_cold_exact']
 assert p['numerical_material_residual_methods_exact_source']
assert not any(p['physical_capacity_certified'] for d in proofs.values() for p in d.values())
check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
native={}
for k in roots:
 jobs=['RealColdReplay'] if k in (793,795,796,798,799,801) else ['HBackendStock','OperationalArcRegressions']
 if k==794:jobs.append('RealColdReplay')
 for job in jobs:
  n,_=check(roots[k],job);assert n['failed']==0 and n['own_process_closed'] and not n['timeout600']
  if job=='OperationalArcRegressions':assert n['passed']==164 and n['macros']==5
  native[f'{k}/{job}']=n
state_path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json'
state=json.loads(state_path.read_bytes());assert not state['active_jobs']
out=repo/'docs/reports/evidence/Coupled6DAdaptivePreEvent_2026-11-25'
assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
for k,r in roots.items():
 files=selected[k]+['Preparation.json','CompilePoint.log','CompilePointClosure.json',
  'PointTraceSourceComposition.json','CompiledSourceReuse.json','CompiledSourceReuseProof.json',
  'docs/regression/CodeUpdate.json','ParentCodeUpdate.json','ParentCompileClosure.json',
  'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls',
  'src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteResponse.cls',
  'src/Materials/CCoupledConcreteReturn.cls','tests/modTestCoupledElasticState.bas',
  'Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1','Run-TestHMixedStock.ps1',
  'Guard-IsolatedOffice.ps1','tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
 # Копируются фактически используемые двоичные входы. Унаследованные старые пробы не являются новым результатом.
 files += ['Input/'+n for n in ('Reverse1.bin','SourceColdRawHistory.bin','ArcCallAllowance.bin',
  'OperationalArcRatio.bin','NativeTolerance.bin','NativeSolver.json','EnvelopeOptions.bin',
  'HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin','MaterialProbeRatio.bin')]
 files += [p.relative_to(r).as_posix() for p in (r/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
 for relative in files:
  p=r/relative
  if p.is_file():
   dest=out/str(k)/relative;dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(p.read_bytes())
helpers=['CandidatePreEventApproach.cls.txt','Stage-TestAdaptivePreEvent.py','Check-TestAdaptivePreEvent.py',
 'Stage-TestPreEventBudget.py','Stage-TestPreEventLimits.py','Stage-TestArcGrowthControl.py',
 'Stage-TestAdaptiveColdReplay.py','Stage-TestReusedAdaptiveColdReplay.py','Stage-TestPostArcAccuracy.py',
 'Check-TestOperationalCold.py','Check-TestOperationalColdReplay.py','Finalize-TestOperationalArc.py',
 'Run-TestOperationalArcRegressions.ps1','Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py',
 'Compile-TestExternalPoint.ps1','Finalize-TestStockYield.py','Audit-TestPrivateCoupledBook.py',
 'Read-TestColdRawCapture.py','Check-TestActualReverseCapture.py','Check-TestPersistentHEvidence.py',
 'Validate-TestHardPrepared.py','Check-TestHardHalfSteps.py',Path(__file__).name]
for n in helpers:(out/n).write_bytes((here/n).read_bytes())
release={n:sha256((repo/n).read_bytes()).hexdigest().upper() for n in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary=dict(proofs=proofs,native=native,release_hashes=release,
 first_cold_operational_arc_admission_confirmed=True,physical_capacity_certified=False,
 inherited_proofs_not_relabelled_as_new=True,all_started_stages_terminal=True,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
 adaptive_pre_event_evidence='docs/reports/evidence/Coupled6DAdaptivePreEvent_2026-11-25',
 previous_goal_turn_classification='progress: bounded pre-event refinement793/794 passes unchanged field and lambda criteria on third halving; exact restore0, no material-history acceptance in operation. First actual cold arc admission794, 67states, all independent equilibrium798. Budget6/3 and mismatched-history negatives795/796 pass. Growth797 yields47 states, independent799; post-arc precision800 yields45 states, independent801. All3 cold runs end NumFail, no higher lambda or physical limit, no speedup claim.164 regression checks each794/797/800 pass; saved/reopened256 actual sources exact. Production unchanged.',
 next_action='Investigate numerical-material return precision or exact constitutive/tangent kink at cold800 endpoint. Its last accepted residual1.148e-10 fits inner1.217e-10, but failed next probe1.961e-9. Tightening post-arc precision did not pass barrier. Local arc alpha finite (~0.0138), so metric blowup unproved. If testing nested numerical tolerances, change outer CDP and scalar return together (prior771/772 evidence), retain physical material,30/96 budgets and all acceptance checks. Need actual cold reproducibility, confirmed physical ultimate, step/mesh/external controls, allLoadPaths and full original deliverables; full long acceptance NOT RUN.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (state_path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('793-801 ARCHIVED',len(pins),'PINS; FIRST COLD ARC ADMISSION; NO PHYSICAL LIMIT; GOAL ACTIVE')
