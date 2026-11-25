"""ДЛЯ ТЕСТОВ. Сохраняет диагностику848–850 и проверенное исправление851, оставляя запущенные852/853 активными."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
from io import BytesIO
import json,tarfile

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={n:next(base.glob(f'Cold{n}*')) for n in (848,849,850,851)}
names={848:'NewLoadingCoordinateDiagnosisProof.json',849:'OutgoingTangentMetricProof.json',850:'OutgoingTangentMetricProof.json',851:'LoadingSecantReparameterizationProof.json'}
proofs={str(n):json.loads((r/names[n]).read_bytes()) for n,r in roots.items()}
assert all(p['coordinate_records'][0]['stage']=='RadiusAndLambda' for p in proofs['848']['cases'][1:])
assert not proofs['851']['cases'][0]['arc_solved'] and all(p['arc_solved'] for p in proofs['851']['cases'][1:])
assert proofs['851']['cases'][1]['independent_field_half_error']<.005
assert all(not p['physical_capacity_certified'] for p in proofs.values())
check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'];native={}
for n,r in roots.items():
    native[str(n)],_=check(r,'RealColdReplay')
    assert native[str(n)]==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
    assert json.loads((r/'CompilePointClosure.json').read_bytes())['State']=='Absent'
native['851_regressions'],_=check(roots[851],'OperationalArcRegressions')
assert native['851_regressions']==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
out=repo/'docs/reports/evidence/Coupled6DLoadingSecantReparameterization_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
common=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
    'CompiledSourceReuse.json','CompiledSourceReuseProof.json','ParentCodeUpdate.json','ParentCompileClosure.json',
    'docs/regression/CodeUpdate.json','Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1','Run-TestRealColdReplay.ps1','Run.ndjson',
    'src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls',
    'src/Solver/CCoupledConstraintEquations.cls','src/Solver/CCoupledMaterialConstraint.cls','src/Solver/CCoupledEnvelopeEquations.cls',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls','src/Materials/CCoupledConcreteReturn.cls',
    'tests/modTestCoupledElasticState.bas','tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
stages={848:'Stage-TestNewLoadingCoordinateDiagnosis.py',849:'Stage-TestOutgoingTangentMetric.py',
    850:'Stage-TestRefinedOutgoingTangentMetric.py',851:'Stage-TestLoadingSecantReparameterization.py'}
inputs=['Reverse1.bin','EnvelopeOptions.bin','NativeTolerance.bin','NativeSolver.json','ArcCallAllowance.bin',
        'OperationalArcRatio.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin']
raw={}
for n,r in roots.items():
    files=common+[names[n],stages[n]]+['Input/'+name for name in inputs]
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
helpers=list(stages.values())+['Check-TestNewLoadingCoordinateDiagnosis.py','Check-TestOutgoingTangentMetric.py',
    'TestUnacceptedMechanics.py','Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py','Check-TestPersistentHEvidence.py',
    'Check-TestActualReverseCapture.py','Read-TestColdRawCapture.py','Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py',
    'Finalize-TestOperationalArc.py','Finalize-TestStockYield.py','Inspect-TestResolutionVBA.py','Audit-TestPrivateCoupledBook.py',
    'Run-TestOperationalArcRegressions.ps1',Path(__file__).name]
for name in sorted(set(helpers)):(out/name).write_bytes((here/name).read_bytes())
release={name:sha256((repo/name).read_bytes()).hexdigest().upper() for name in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary=dict(proofs=proofs,native=native,release_hashes=release,raw_files=len(raw),raw_archive=archive.name,
    raw_bytes=sum(map(len,raw.values())),material_laws_unchanged=True,original_locality_and_pair_criteria_unchanged=True,
    automatic_sequence852_and_cold853_not_part_of_this_completed_checkpoint=True,
    not_clean_performance_benchmark=True,physical_capacity_certified=False,full_acceptance_completed=False,
    whole_goal_completed=False,goal_active=True)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
    pause_request_revoked=True,pause_after_current_stage=False,
    latest_verified_loading_secant_source=str(roots[851]),loading_secant_reparameterization_evidence=out.relative_to(repo).as_posix(),
    loading_secant_local_summary=dict(source_states=86,point=2,side=-1,ratios_tested=[1,.5,.25,.125,.0625,.03125],
        largest_ratio_still_rejected=True,five_smaller_ratios_verified=True,half_error_at_ratio_half=proofs['851']['cases'][1]['independent_field_half_error'],
        all_original_radii_and_pair_criteria_preserved=True,regressions=164,physical_capacity_certified=False),
    previous_goal_turn_classification='progress:848 pinpoints next86 failure as radius rejection AFTER converged scalar/force roots; repeatedcutbacks retain3.1768radiusratio.849ordinarytangentsunstable;850existingrefinedbranch at1/64incoming nearlyexactmatchesobserved outgoingdirection(cos0.9999999999458902,gain19.0607).851workingCArc computes its own unaccepted equilibriumguide, requiresbitexactzeroiterationrestoration, normalizes its physical displacementtoORIGINALrequestedspacing,andretargets scalar pairWITHOUTchanging1.5/3radii,.005fullhalf,lambda/force/resolution/direction/steelcriteria,30sharedphaseor96callbudget. Five smaller ratios nowverified;largeststillrejected(error.2899). Actual256/2imports254othersexact;164regressionspassed.852automaticfrom70and853coldcurrentlyseparateactivechecks,notcompletion.',
    next_action='Continue full original goal;DO NOT PAUSE without newexplicitrequest. Current livejobs in active_jobs are852warm and853cold (verifyactualhandles beforeassuminglive). Local851fixverifiedfrom86:5of6ratiossolve,thelargestcorrectlyrejects. Kernel sourceCold851LoadingSecantReparameterization_20261125;CArc newTryLoadingDeparturePair exactpostconditions,verifiedunacceptedhalfguide thenkinematicallynormalizedprediction;allguide/retry/restorecalls counted,original30phase/96and1.5/3radii preserved,phaseguideszeroiterations. Newton/materials exact848 except opt-in readonly tracing;its working methods exact846 asproved. First run Check-TestAutomaticLoadingWalk.py --root D:/CodexScratch/RC_Section_NDM_TQ/Cold852AutomaticLoadingSecantWalk_20261125 --continued-prefix --auto-h-guide-prefix --loading-secant-prefix afterhandle42373terminal. NewGuidev2parser andindependentguideKstress supportprepared. Then Check-TestAutomaticHGuideCold.py --loading-secant afterhandle73914terminal;outputLoadingSecantColdProof.json. If coldhistorydiffers832, performexactnativefrozenstressreplaybeforeclaimindepcold verification (Stage-TestOperationalColdReplay orStage-TestAdaptiveColdReplay pattern;pinallactualsourceandcalibration). Cold847previousstill49,max1.1058869810286156,NumFail;donotclaimnewcoldimprovementuntilactual853. 850slowcompilationeventuallycompletedwithoutintervention,ownclosed;noabort/noUIinput. IndependentrefinedtangentdiagnosisusesexistingTryReadBranchLinearization only;no tangentpolicy/materialpromotion. After automatic/cold results resolve anynewblockingtransition,repeatdirectedaffectedchecks,archiveandpush. FullphysicalcriterionONQ010/012,step/mesh/external/allLoadPaths/hoop/UI/CAD/cracks/stability/performance/release/fullscopedacceptanceremainopen. Main/releaseunchanged;priorverifiedGitHub28a2f4f43c91dd050797ac521ca49a0938caac0d.')
state['terminal_sessions_this_stage']=sorted(set(state.get('terminal_sessions_this_stage',[])+[69179,5187,82906,26175,83496,68709,39642,75997,92748,91139,47981,12254,91793,1801,78836]))
encoded=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(encoded)
print('LOADING SECANT ARCHIVED:',len(raw),'RAW FILES;',archive.stat().st_size,'BYTES;164REGRESSIONS;LIVE JOBS PRESERVED',len(state.get('active_jobs',[])))
