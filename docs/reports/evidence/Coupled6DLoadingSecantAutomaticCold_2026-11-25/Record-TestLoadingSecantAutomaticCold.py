"""ДЛЯ ТЕСТОВ. Сохраняет автоматические852/853 и физический аудит малости прогресса; число точек не подменяет предел."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
from io import BytesIO
import json,tarfile

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={852:base/'Cold852AutomaticLoadingSecantWalk_20261125',853:base/'Cold853LoadingSecantCold_20261125'}
warm=json.loads((roots[852]/'AutomaticLoadingWalkProof.json').read_bytes())
cold=json.loads((roots[853]/'LoadingSecantColdProof.json').read_bytes())
audit=json.loads((roots[852]/'ContinuationProgressAudit.json').read_bytes())
assert warm['new_accepted_equilibria']==22 and warm['total_calls']==94 and warm['all_new_H_strictly_growing']
assert cold['accepted_states']==49 and cold['cold832_fields_fractions_lambdas_bit_identical'] and cold['cold']['capacity']=='NumFail'
assert audit['cold_budget']['stop_by_global_probe_limit'] and audit['cold_budget']['last_arc_restore_succeeded']
check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'];native={}
for n,r in roots.items():
    native[str(n)],_=check(r,'RealColdReplay' if n==852 else 'HBackendStock')
    assert native[str(n)]['failed']==0 and not native[str(n)]['timeout600'] and native[str(n)]['own_process_closed']
out=repo/'docs/reports/evidence/Coupled6DLoadingSecantAutomaticCold_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
common=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
    'CompiledSourceReuse.json','CompiledSourceReuseProof.json','ParentCodeUpdate.json','ParentCompileClosure.json',
    'docs/regression/CodeUpdate.json','Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1','Run-TestRealColdReplay.ps1',
    'Run-TestHMixedStock.ps1','Run.ndjson','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledNewtonSolver.cls',
    'src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteResponse.cls',
    'src/Materials/CCoupledReinforcedResponse.cls','src/Materials/CCoupledConcreteReturn.cls','tests/modTestCoupledElasticState.bas']
stages={852:'Stage-TestAutomaticLoadingSecantWalk.py',853:'Stage-TestLoadingSecantCold.py'}
proof_names={852:['AutomaticLoadingWalkProof.json','ContinuationProgressAudit.json'],853:['LoadingSecantColdProof.json']}
inputs=['Reverse1.bin','EnvelopeOptions.bin','NativeTolerance.bin','NativeSolver.json','ArcCallAllowance.bin',
        'OperationalArcRatio.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin']
raw={}
for n,r in roots.items():
    files=common+proof_names[n]+[stages[n]]+['Input/'+name for name in inputs]
    files += [p.relative_to(r).as_posix() for p in (r/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        p=r/name
        if p.is_file():raw[f'{n}/{name}']=p.read_bytes()
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=6) as z:
    content=dict(raw);content['RawFiles.json']=(json.dumps(pins,indent=2,sort_keys=True)+'\n').encode()
    for name,data in content.items():
        info=tarfile.TarInfo(name);info.size=len(data);z.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as z:
    for name,digest in pins.items():assert sha256(z.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=list(stages.values())+['Check-TestAutomaticLoadingWalk.py','Check-TestAutomaticHGuideCold.py',
    'Check-TestContinuationProgress.py','TestUnacceptedMechanics.py','Stage-TestEnergyNumerics.py',
    'Check-TestPersistentHEvidence.py','Check-TestActualReverseCapture.py','Read-TestColdRawCapture.py',
    'Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py','Inspect-TestResolutionVBA.py',Path(__file__).name]
for name in sorted(set(helpers)):(out/name).write_bytes((here/name).read_bytes())
summary=dict(warm=warm,cold=cold,progress_audit=audit,native=native,raw_archive=archive.name,raw_files=len(raw),
    raw_bytes=sum(map(len,raw.values())),cold_capacity_result_unchanged=True,physical_progress_small_despite_more_states=True,
    not_clean_performance_benchmark=True,physical_capacity_certified=False,whole_goal_completed=False,goal_active=True)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
    pause_request_revoked=True,pause_after_current_stage=False,active_exec_session=None,active_jobs=[],all_started_stages_terminal=True,
    loading_secant_automatic_cold_evidence=out.relative_to(repo).as_posix(),
    latest_actual_cold_source=str(roots[853]),latest_verified_automatic_Hwalk_source=str(roots[852]),
    latest_automatic_Hwalk_summary=dict(accepted=22,total_states=92,calls=94,attempts=30,last_lambda=1.1000749304461266,
        all_new_H_strictly_growing=True,loading_guides=2,physical_progress_small=True,physical_capacity_certified=False),
    latest_actual_cold_summary={k:cold[k] for k in ('cold','accepted_states','maximum_accepted_lambda','last_accepted_lambda','cold832_fields_fractions_lambdas_bit_identical')},
    latest_continuation_physical_progress_audit=audit,
    previous_goal_turn_classification='progress:local851verifiedguide reparameterization preservesoriginalradii/force/fullhalf/lambda/30and96limits,5of6ratiosnowpass,largeststillrejects.164regressionspass.852automaticfrom70crossesprevious86barrier,reaches92(+22total)in94calls;allroots/halves/guides/forces/gauges/steel/Hindepverified.853actualcoldsaved/reopened256BUTexactold49states/max1.1058869810286156/NumFail;NO COLD CAPACITY IMPROVEMENT. Usercorrectlycriticizedcountgrowthwithoutlambda;physicalauditshows86to92lambda+9.302e-5,concreteprincipalmax~.00044064andlongsteel~8.35e-5nearlyunchanged. Coldlastarcvalidbutglobal96exhausted;at42states73probes,then7states23moreprobes. ThereforecountgrowthNOTevidenceofnearultimate;prioritizecoldbudgetallocationandphysicalprogress.',
    next_action='Continue full original goal;DO NOT PAUSE absent newexplicitrequest. All nativejobs852/853terminal,ownclosed. Candidate851 verifiedlocalfix,852warm92onlysmallphysicalprogress,853coldEXACT83249/NumFail. Read latest_continuation_physical_progress_audit and reports beforeclaimingprogress. Immediateprioritycoldfromzero: inspectcontrollerProbe/globalallocation,notmorewarmmicrostates. Lastcold96stopisbudget AFTERsuccessfularc (31arccalls8accepts5retained);priorcheckpoint42states73probes,final49/96. SourceconstantPROBE_LIMIT96inCCoupledDeformationContinuation,notaphysicalcriterionoruserTZlimit. Need choosejustifiedearlyarc/recoveryafterexactfailedmeasure toavoidrepeatedwrongcoordinate (cachedfailedrevision/projectionalreadyexists),preserve831retained/proactivecontextguards andnewCArcH/steel/verifiedguidechecks. Alternatively a clearly separated extended-budget cold diagnostic (e.g.192,predeclaredtimeout/resourceguard) can revealwhetherpostcrackphysicsprogresses;DONOTpromotehigherbudgetasstable/fasterorcapacityproof. Preserveoriginal96negativecontrol. Measureprogressinprincipalstrains/steel/degradation/globalphysicalcriterion,notstatecountoroneλmaximum. Do not blindly relaxradii orforce/half thresholds. 848diagnosisradiusratio3.1768;850refinedbranch at1/64predictsactualoutgoinggain19.06butlargeincomingprobeisunstable;851usesverifiedsecantguideandretargetingwithORIGINAL1.5/3radii. WorkingCArc851+Newton848opt-inreadonlytrace+Deform845(pinned831withnextstepcap1) arecurrentcandidate. 852failsnextfaceatstate92,4attemptsratios1,.5,.25,.125using16,16,15,15primary;secondaryfuturediagnosiscanreusecompiled851TestDiagnoseNewLoadingCoordinate withnewInput92 toavoidcompile (verifyactualprefix86matchbeforeTestClone). FulloriginalphysicalONQ010/012/mesh/step/external/allLoadPaths/hoop/UI/CAD/cracks/stability/performance/release/fullscopedacceptanceremainopen. NeedGitcheckpointandpushnow;lastverifiedremote28a2f4f43c91dd050797ac521ca49a0938caac0d. Newlocalreports/evidence848–853ready;no main/releasechange.')
state['terminal_sessions_this_stage']=sorted(set(state.get('terminal_sessions_this_stage',[])+[42373,73914,9644]))
encoded=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(encoded)
print('AUTOMATIC/COLD/PHYSICAL AUDIT ARCHIVED:',len(raw),'FILES;',archive.stat().st_size,'BYTES;COLD UNCHANGED;GOAL ACTIVE')
