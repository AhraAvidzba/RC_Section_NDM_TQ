"""ДЛЯ ТЕСТОВ. Сохраняет автоматическое H-направление в рабочей дуге, независимые бюджетные контроли и исходный неудачный диагностический прогон."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
from io import BytesIO
import json,tarfile

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={843:base/'Cold843AutomaticHGuideArc_20261125',844:base/'Cold844AutomaticHGuideArc_20261125'}
proof=json.loads((roots[844]/'AutomaticHGuideArcProof.json').read_bytes())
assert proof['allowance5_unloading_root_not_admitted_with_two_calls']
assert proof['allowance6_and96_return_bit_identical_event']
assert len(proof['successful_cases'])==2 and not proof['physical_capacity_certified']
assert all(p['calls']==6 and p['pair_primary']==18 and p['recovery_primary']==8 and p['total_primary']==26 for p in proof['successful_cases'])
check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'];native={}
for label,passed in [('RealColdReplay',1),('OperationalArcRegressions',164)]:
    value,_=check(roots[844],label);assert value['passed']==passed and value['failed']==0 and value['own_process_closed'] and not value['timeout600']
    native[label]=value
failed_log=(roots[843]/'Native/RealColdReplay.txt').read_text('utf-8-sig')
assert 'RUNTIME ERROR:' in failed_log and 'WATCHDOG_COMPLETED: exit=1; source unchanged=True' in failed_log
failed_closure=json.loads((roots[843]/'Native/RealColdReplay.txt.ProcessClosure.json').read_text('utf-8-sig'))
assert 'Absent' in json.dumps(failed_closure)
for n,r in roots.items():
    assert json.loads((r/'CompilePointClosure.json').read_bytes())['State']=='Absent'
    audit=json.loads((r/'PointTraceSourceComposition.json').read_bytes());assert audit['components']==256
    assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==254
out=repo/'docs/reports/evidence/Coupled6DAutomaticHGuideArc_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
common=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
    'CompiledSourceReuse.json','CompiledSourceReuseProof.json','ParentCodeUpdate.json','ParentCompileClosure.json',
    'docs/regression/CodeUpdate.json','Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1','Run-TestRealColdReplay.ps1',
    'Stage-TestAutomaticHGuideArc.py','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledNewtonSolver.cls',
    'src/Solver/CCoupledDeformationContinuation.cls','src/Solver/CCoupledMaterialConstraint.cls','src/Solver/CCoupledEnvelopeEquations.cls',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls','src/Materials/CCoupledConcreteReturn.cls',
    'tests/modTestCoupledElasticState.bas','tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
inputs=['Reverse1.bin','EnvelopeOptions.bin','NativeTolerance.bin','NativeSolver.json','ArcCallAllowance.bin',
    'OperationalArcRatio.bin','SourceColdRawHistory.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin']
raw={}
for n,r in roots.items():
    files=common+['Input/'+name for name in inputs]
    if n==844:files+=['AutomaticHGuideArcProof.json']
    files += [p.relative_to(r).as_posix() for p in (r/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        p=r/name
        if p.is_file():raw[f'{n}/{name}']=p.read_bytes()
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()}
archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=6) as z:
    content=dict(raw);content['RawFiles.json']=(json.dumps(pins,indent=2,sort_keys=True)+'\n').encode('utf-8')
    for name,data in content.items():
        info=tarfile.TarInfo(name);info.size=len(data);z.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as z:
    for name,digest in pins.items():assert sha256(z.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=['Stage-TestAutomaticHGuideArc.py','Check-TestAutomaticHGuideArc.py','TestUnacceptedMechanics.py',
    'Stage-TestEnergyNumerics.py','Check-TestPersistentHEvidence.py','Check-TestActualReverseCapture.py',
    'Read-TestColdRawCapture.py','Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py',
    'Finalize-TestOperationalArc.py','Finalize-TestStockYield.py','Inspect-TestResolutionVBA.py','Audit-TestPrivateCoupledBook.py',
    'Run-TestOperationalArcRegressions.ps1',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
release={name:sha256((repo/name).read_bytes()).hexdigest().upper() for name in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary=dict(automatic_guide=proof,native=native,release_hashes=release,raw_archive=archive.name,raw_files=len(raw),
    raw_bytes=sum(map(len,raw.values())),failed843_preserved=True,
    failed843_reason='Only full step converged; original classification required verified pair; test output read absent half root and aborted.',
    corrected844_converged_full_or_half_can_trigger_H_guide_without_admitting_unverified_pair=True,
    material_and_Newton_equations_unchanged=True,original_arc_pair_and_event_acceptance_math_unchanged=True,
    H_guide_only_locator_not_accepted_equilibrium=True,
    active_steel_plasticity_guard_and_general_controller_validation_still_required=True,
    step_ratio_policy_after_new_loading_boundary_still_to_validate_in_retained_controller=True,
    not_cold_capacity_search=True,not_clean_performance_benchmark=True,
    physical_capacity_certified=False,full_acceptance_completed=False,whole_goal_completed=False,goal_active=True)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
event=proof['successful_cases'][0]
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
    pause_request_revoked=True,pause_after_current_stage=False,active_exec_session=None,active_jobs=[],all_started_stages_terminal=True,
    automatic_H_guide_evidence=out.relative_to(repo).as_posix(),latest_verified_auto_H_guide_source=str(roots[844]),
    automatic_H_guide_summary=dict(calls=6,pair_primary=18,recovery_primary=8,total_primary=26,
        event_lambda=event['event_lambda'],event_point=event['event_point'],event_side=event['event_side'],
        allowance6_and96_bit_identical=True,allowance5_not_admitted=True,regression_passed=164,
        no_precomputed_future_field=True,physical_capacity_certified=False),
    previous_goal_turn_classification='progress:working CArc now detects a converged unloading full/half candidate and computes its own H guide,then verifies material-event and half pair within shared30 recovery phase.843 failed diagnostic/pair-eligibility case preserved and fixed.844 original70 test runs with futureHfieldinputabsent;allowance5 uses2calls and doesnotadmitunloadingroot;allowances6/96 use6calls,18initialpair+8sharedrecovery=26primary,andreturnbitidenticalloadingeventlambda'+str(event['event_lambda'])+'.Allforces/gauges/three-stationsteel/H/phasebudgets/radius/exactrestore independently verified;164directedregressionspassed.Cold/generalsteel/capacityproofnotcomplete.',
    next_action='Continue full original active goal. Working candidate844 now computes H guide itself in CArc.TryStep (noInput/HMaterialGuide.bin), triggers on convergedSF evenifSHfailed, returns verified25/-1 eventatlambda'+str(event['event_lambda'])+';code +native5/6/96 budget controls +164regressionsverified. Before general promotion, add a READONLY steel-plastic-activity guard: frozen concrete H plus declininglambda must not automatically divert a valid steel-plastic/softening branch. Reuse reinforced.ReadSteelBranch and all3stations;materialsunchanged. Test elastic vsplastic longitudinal/transverse andnoSteel cases. Also owner next-ratio policy after newloadingphi face:842 required ratios1,.5,.25,.125 before successful kind4; normal warm driver after kind2 suggests2 and wouldexhaust4attempts before.125. Expose/reset nextstep bound1 afternewloadingphi eventonly (not spectra) in both retained831controller and warm driver, preserving force/fullhalf/budgets. Then run automatic sequence from original70 with newguide, throughfaceandoutgoing withoutprescribed intermediatefield, monitorH and continue exact842active79 prefix tofindphysicalend. Integrate current844kernel with actual831/832gatedcoldcontroller andrun nativecold,targetedregressions;cold832still49/max1.1058869810286156/NumFail96. NewH guide,event,half,restore share30phase andcountallcalls;guideunadmitted,eventlocal≤3originalradii. Standalone843failurepartialrawpreserved;noPARSEpassclaim. Need loading/steel distinction,ONQ010/012physicalcriterion,step/mesh/external andfull originalLoadPaths/hoop/UI/CAD/cracks/stability/performance/actualreleasebook/full scopedacceptance. Main/releaseunchanged. Prior successfulGitHubcheckpoint51dca39;pushlatestwithHTTPSHTTP/2,postBuffer536870912,emptyExpectandverifyactualremote.')
state['terminal_sessions_this_stage']=sorted(set(state.get('terminal_sessions_this_stage',[])+[99214,69535,69420,17224,36477,50306,23313]))
encoded=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(encoded)
print('AUTOMATIC H GUIDE ARCHIVED:',len(raw),'RAW FILES;',archive.stat().st_size,'BYTE TAR.XZ;164REGRESSIONSPASS;GOAL ACTIVE')
