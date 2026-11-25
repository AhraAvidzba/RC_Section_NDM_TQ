"""ДЛЯ ТЕСТОВ. Сохраняет проверки845–847, исходные сертификаты и честный холодный результат; цель остаётся активной."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
from io import BytesIO
import json,tarfile

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={n:next(base.glob(f'Cold{n}*')) for n in (845,846,847)}
names={845:'SteelActivityArcPolicyProof.json',846:'AutomaticLoadingWalkProof.json',847:'AutomaticHGuideColdProof.json'}
proofs={str(n):json.loads((r/names[n]).read_bytes()) for n,r in roots.items()}
assert proofs['845']['individual_station_bar_branches_verified']==702
assert proofs['846']['new_accepted_equilibria']==16 and proofs['846']['all_new_H_strictly_growing']
assert all(not p['physical_capacity_certified'] for p in proofs.values())
check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'];native={}
for n,r in roots.items():
    native[str(n)],_=check(r,'HBackendStock' if n==847 else 'RealColdReplay')
    assert native[str(n)]['failed']==0 and native[str(n)]['own_process_closed'] and not native[str(n)]['timeout600']
native['845_regressions'],_=check(roots[845],'OperationalArcRegressions')
assert native['845_regressions']==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
out=repo/'docs/reports/evidence/Coupled6DSteelActivityHGuideWalk_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
common=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
    'CompiledSourceReuse.json','CompiledSourceReuseProof.json','ParentCodeUpdate.json','ParentCompileClosure.json',
    'docs/regression/CodeUpdate.json','Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1','Run-TestRealColdReplay.ps1',
    'Run-TestHMixedStock.ps1','Run.ndjson','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledNewtonSolver.cls',
    'src/Solver/CCoupledDeformationContinuation.cls','src/Solver/CCoupledMaterialConstraint.cls','src/Solver/CCoupledEnvelopeEquations.cls',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls','src/Materials/CCoupledConcreteReturn.cls',
    'tests/modTestCoupledElasticState.bas','tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
stages={845:'Stage-TestSteelActivityAndArcPolicy.py',846:'Stage-TestAutomaticHGuideWalk.py',847:'Stage-TestAutomaticHGuideCold.py'}
raw={}
for n,r in roots.items():
    files=common+[names[n],stages[n]]
    files += [p.relative_to(r).as_posix() for p in (r/'Input').iterdir() if p.is_file()]
    files += [p.relative_to(r).as_posix() for p in (r/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        p=r/name
        if p.is_file():raw[f'{n}/{name}']=p.read_bytes()
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()}
archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=6) as z:
    content=dict(raw);content['RawFiles.json']=(json.dumps(pins,indent=2,sort_keys=True)+'\n').encode()
    for name,data in content.items():
        info=tarfile.TarInfo(name);info.size=len(data);z.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as z:
    for name,digest in pins.items():assert sha256(z.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=list(stages.values())+['Check-TestSteelActivityAndArcPolicy.py','Check-TestAutomaticLoadingWalk.py',
    'Check-TestAutomaticHGuideCold.py','TestUnacceptedMechanics.py','Stage-TestEnergyNumerics.py',
    'Check-TestPersistentHEvidence.py','Check-TestActualReverseCapture.py','Read-TestColdRawCapture.py',
    'Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py','Finalize-TestOperationalArc.py','Finalize-TestStockYield.py',
    'Inspect-TestResolutionVBA.py','Audit-TestPrivateCoupledBook.py','Run-TestOperationalArcRegressions.ps1',Path(__file__).name]
for name in sorted(set(helpers)):(out/name).write_bytes((here/name).read_bytes())
release={name:sha256((repo/name).read_bytes()).hexdigest().upper() for name in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary=dict(proofs=proofs,native=native,release_hashes=release,raw_files=len(raw),raw_archive=archive.name,
    raw_bytes=sum(map(len,raw.values())),material_laws_unchanged=True,
    next_loading_boundary_remains_numerically_unresolved=True,cold_result_not_reclassified_as_physical_failure=True,
    not_clean_performance_benchmark=True,physical_capacity_certified=False,full_acceptance_completed=False,
    whole_goal_completed=False,goal_active=True)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
    pause_request_revoked=True,pause_after_current_stage=False,active_exec_session=None,active_jobs=[],all_started_stages_terminal=True,
    latest_steel_guard_and_automatic_Hwalk_evidence=out.relative_to(repo).as_posix(),latest_candidate_source=str(roots[845]),
    latest_verified_automatic_Hwalk_source=str(roots[846]),latest_actual_cold_source=str(roots[847]),
    latest_automatic_Hwalk_summary=dict(accepted=16,total_states=86,calls=61,attempts=23,last_lambda=1.099981906664804,
        all_new_H_strictly_growing=True,next_phi_boundary_outgoing_not_verified=True,physical_capacity_certified=False),
    latest_actual_cold_summary={k:proofs['847'][k] for k in ('cold','accepted_states','maximum_accepted_lambda','last_accepted_lambda','cold832_fields_fractions_lambdas_bit_identical')},
    previous_goal_turn_classification='progress:readonly all-station longitudinal/hoop steel plastic guard passed13kinematic fields,702station/bar branch checks,real plastic axialNDS with unload/reverse,and no-steel case. Actual256candidatecompiled4imports252others.164targetedregressionspassed. Newphi-face nextstep cap1 integrated into pinned831owner. Native846 from exact70 WITHOUT predetermined Hroot/face computed its own guide,event and outgoingkind4;16new equilibria with strict full/half/sixforce/steel/gauge checks and growingH reached86/lambda1.099981906664804,then4outgoingattempts failedat nextphi face. Cold847 honest saved/reopened256 outcome recorded separately;no physicalultimateclaim.',
    next_action='Continue full original active goal;DO NOT PAUSE without new explicit request. Candidate845 plus native846 automaticHguide works from70 through16newstates,allHgrow,endsat new phi face sourcecount86/lambda1.099981906664804 (attempt19);attempts20..23 ratios1,.5,.25,.125fail,61callsused (35available). Diagnose this next recent-loading departure using exact846 prefix86/event packet and readable point/station/yield state,not older838unloadingbranch. Inspect why full/half or material branch fails;consider broader adaptivecutback or consistent outgoing tangent within unchanged30phase/96call budget,not numericalfailureasphysicalfailure. A diagnostic continuation may test smaller outgoing steps, then integrate justified common policy into CArc owner+warmdriver andrepeatactualcold. Preserve original831retained/proactive contextguard tests,164regressions,exact full/half/radius/force/steel checks. Steel guard usesexistingReadSteelBranch atall3stations,realplasticfixture154938.5764N/401.5MPa;no materiallaw changes. Cold847 summaryinlatest_actual_cold_summary;warmprogressdoesnotprove coldimprovement. Need ONQ010/012physicalcriterion,step/mesh/external/allLoadPaths/hoop/UI/CAD/cracks/stability/performance/actualreleasebook/full scopedacceptance. Main/releaseunchanged. PushcandidatebranchHTTPSHTTP2andverifyactualremote;lastverified90784a1.')
state['terminal_sessions_this_stage']=sorted(set(state.get('terminal_sessions_this_stage',[])+[43884,4388,79320,97396,98340,15306,13853,49065]))
encoded=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(encoded)
print('STEEL/HWALK/COLD ARCHIVED:',len(raw),'RAW FILES;',archive.stat().st_size,'BYTES;164REGRESSIONS;GOAL ACTIVE')
