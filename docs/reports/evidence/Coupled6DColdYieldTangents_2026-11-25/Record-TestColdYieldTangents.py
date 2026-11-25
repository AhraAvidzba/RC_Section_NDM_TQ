"""ДЛЯ ТЕСТОВ. Сохраняет настоящие причины отказа скалярных проб и проверенные исходящие касательные, не объявляя прогноз равновесием."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from io import BytesIO
from datetime import datetime
import json,tarfile,subprocess

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={863:base/'Cold863Cold19YieldCoordinateCapture_20261125',864:base/'Cold864Cold19BranchTangents_20261125'}
proofs={863:'ColdYieldCoordinateDiagnosisProof.json',864:'ColdBranchTangentsProof.json'}
results={n:json.loads((roots[n]/name).read_bytes()) for n,name in proofs.items()}
assert len(results[863]['cases'])==6 and not any(c['arc_solved'] for c in results[863]['cases'])
records=[r for c in results[863]['cases'] for r in c['coordinate_records']];assert len(records)==3
assert all(r['stage']=='EquilibriumAndGauge' and r['mechanics']['dominant_equation']==55 for r in records)
assert all(abs(r['scalar_residual_MPa'])<1e-11 and r['mechanics']['weak_norm_MPa']>1e-3 for r in records)
assert len(results[864]['tangents'])==4 and all(t['available'] and t['rank_1e_minus12']==60 for t in results[864]['tangents'])
assert all(t['same_length_lambda_increment']<0 and t['cosine_with_incoming_secant']<-.99 for t in results[864]['tangents'][1:])
native={};check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
for n,r in roots.items():
    native[n],_=check(r,'RealColdReplay');assert native[n]['failed']==0 and native[n]['own_process_closed'] and not native[n]['timeout600']
out=repo/'docs/reports/evidence/Coupled6DColdYieldTangents_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
common=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
 'CompiledSourceReuse.json','CompiledSourceReuseProof.json','ParentCodeUpdate.json','ParentCompileClosure.json',
 'docs/regression/CodeUpdate.json','Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1','Run-TestRealColdReplay.ps1','Run.ndjson',
 'src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls',
 'src/Solver/CCoupledElasticState.cls','src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls',
 'src/Materials/CCoupledConcreteReturn.cls','tests/modTestCoupledElasticState.bas','tools/build_workbook/Run-Audit02ModeTests.ps1']
inputs=['Reverse1.bin','ArcScanFields.bin','EnvelopeOptions.bin','NativeTolerance.bin','NativeSolver.json','ArcCallAllowance.bin',
 'OperationalArcRatio.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin']
extra=['SourceColdStressProof.json','SourceColdContinuityProof.json','SourceArcRestartProof.json']
stages={863:'Stage-TestColdYieldCoordinateCapture.py',864:'Stage-TestColdBranchTangents.py'}
raw={}
for n,r in roots.items():
    files=common+extra+[proofs[n],stages[n],'Align-TestShortEntryInsertion.py']+['Input/'+name for name in inputs]
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
helpers=list(stages.values())+['Check-TestNewLoadingCoordinateDiagnosis.py','Check-TestColdBranchTangents.py',
 'Align-TestShortEntryInsertion.py','Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py',
 'Check-TestPersistentHEvidence.py','Check-TestActualReverseCapture.py','Read-TestColdRawCapture.py',
 'Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py',Path(__file__).name]
for name in sorted(set(helpers)):(out/name).write_bytes((here/name).read_bytes())
summary=dict(results={str(n):v for n,v in results.items()},native={str(n):v for n,v in native.items()},
 raw_archive=archive.name,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),
 actual_global_equilibrium_failure_not_only_locality=True,dominant_failed_equation55_gradientMy=True,
 consistent_outgoing_tangent_predicts_lambda_decline=True,tangent_predictors_not_NDS_or_capacity_results=True,
 first864_source_audit_failed_entry_order_only=True,prepared_source_aligned_to_VBE_insertion_without_method_or_book_change=True,
 repeated_strict_actual256_audit_passed=True,no_working_or_release_promotion=True,
 physical_capacity_certified=False,whole_goal_completed=False,goal_active=True)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
current=subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
 pause_request_revoked=True,pause_after_current_stage=False,latest_cold_yield_tangent_evidence=out.relative_to(repo).as_posix(),
 latest_cold19_outgoing_tangent_proof=results[864],latest_cold19_failed_yield_summary=dict(cases=6,scalar_records=3,dominant_equation=55,force_failure_not_radius_only=True),
 verified_git_delivery=dict(commit=current,remote_branch='codex/coupled-6d',remote_ref_independently_verified=True,
 successful_transport='HTTPS HTTP/2 buffered push; independent dry-run receive-pack handshake returned up to date'),
 next_action='Continue original goal;DO NOT PAUSE. All numericjobs858–864terminal/ownclosed;no physicalcapacity. Completed858–862archivedandpushed80a6ff5,independentpush--dry-run--porcelainconfirmedup-to-date (lsremotepathpreviouslytimedout). Current863/864archiveprimaryfindingsreadycheckpointpush. Criticalnewtangent:atcold19secantpredictspositiveDL+.00051258,ordinaryKatzeroelasticbranchpredicts+.02808/cos+.509;threeFORWARDrefinedKprobes1/1024,1/256,1/64givecos-.9919forpositive-loadderivative;orientedforwardfieldrequiresNEGATIVEdLambda[-.000224538,-.000230995,-.000256905]atsamephysicalstep. Next FIRST testnewfull/halfcontrolledpairsusingorientednativeKdirectionwithnegativepredictedlambda,notnewfuturefield. Useexistingtrue19materialhistory,ratio.125and.0625belowfirstnewtangentyieldfraction~.1663(point1side-1),alloriginalforce/gauge/steel/H/direction/locality/fieldhalf.005/lambdahalf/30pairbudget/postrestore0iter+samefield+lambda;no branchacceptbeforeproof. CanappendshortPublicTestentrytooriginalmoduleusingEntryPointNames(no8kmodule replacement);newentryreads Input/ColdTangentPredictors.bin fromactualColdBranchTangentsProof andcalls existingTryDeformationStep/SelectSpatialDirectionControl toreplicateordinaryarc pair,thenexportsunacceptedfull/halfrootsforindependentmechanics/H. Choosefirstprobe1/1024;comparative1/256optional. Computealpha medianoflast3physicalDistance/absdLambda asCArc,controlsfromtangentv ratherthanfakechangingprevioushistory. KeepgaugedofsseedunchangedlikeexistingCArc. IfpairsPASS thenimplementproperworkingNewtonhelperforbranch-tangentpredictor+OO CArc fallbackinoriginalrecovery30/global96budget;dontcallTest-namedmethodsfromworkingcode(use/refactorcommonprivatehybridsetup). Preserveexistingrecentyield/H/spectralpaths andallpostchecks,then164targetedtestsandTRUEcoldverification. Do notclaimtangentisrootorloadpeakcapacity. 863all6arcsfalse;first3point15/-1scalarphi~4e-14butweak.001304MPa(eq55gradientMy),sixforce~1.2e-6,radiusratios.21/.42/.85;smallratio1/32halferror.003326passesfieldbutstillfalse,inspectlambda/localityratherthanblindcutbacks. 862zero postcommitstress/weakdrift0 andlast19strictroot5.51e-11excludedaccuracyfloor. Fulloriginalscope/ONQ010012/meshstep/external/allLoadPaths/hoop/UI/CAD/cracks/stability/10s/release/fullscopedacceptanceOPEN. PurebendingCDP-SPexplicitlydeferred.')
encoded=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(encoded)
print('COLD YIELD/TANGENT EVIDENCE ARCHIVED',len(raw),'FILES',archive.stat().st_size,'BYTES;PREDICTOR NOT ROOT;GOAL ACTIVE')
