"""ДЛЯ ТЕСТОВ. Сохраняет строгие касательные корни и обоснование постоянной совместной метрики без объявления физического предела."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from io import BytesIO
from datetime import datetime
import json,tarfile

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={865:base/'Cold865Cold19TangentPairs_20261125',866:base/'Cold866Cold19ReferenceMetricPairs_20261126'}
results={n:json.loads((r/'ColdTangentPairsProof.json').read_bytes()) for n,r in roots.items()}
for n,proof in results.items():
    assert proof['native']==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
    assert len(proof['cases'])==3 and all(c['half_solved'] and c['full_solved'] and c['pair_iterations']==4 for c in proof['cases'])
    assert all(c['material_H_advances'] and c['restoration_zero_iterations_same_lambda_bit_identical'] for c in proof['cases'])
    assert proof['true_source19_history_and_final_restored_csv_bit_identical']
assert not any(c['qualifies_original_arc_criteria'] for c in results[865]['cases'])
assert [c['qualifies_joint_metric_diagnostic'] for c in results[866]['cases']]==[False,True,True]
counter=json.loads((here/'ArcMetricCounterexample.json').read_bytes());assert len(counter['cases'])==12
out=repo/'docs/reports/evidence/Coupled6DColdTangentMetricPairs_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
common=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
 'CompiledSourceReuse.json','CompiledSourceReuseProof.json','ParentCodeUpdate.json','ParentCompileClosure.json',
 'docs/regression/CodeUpdate.json','Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1','Run-TestRealColdReplay.ps1','Run.ndjson',
 'src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls',
 'src/Solver/CCoupledElasticState.cls','src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls',
 'src/Materials/CCoupledConcreteReturn.cls','tests/modTestCoupledElasticState.bas','tools/build_workbook/Run-Audit02ModeTests.ps1',
 'SourceColdTangentProof.json','CandidateTestColdTangentPairs.bas.txt','Stage-TestColdTangentPairs.py','ColdTangentPairsProof.json',
 'Align-TestEntryTrailingBlank.py','PreparedTestBeforeBlankCorrection.bas','SupersededFormattingDiagnostic.json']
inputs=['Reverse1.bin','ColdTangentPredictors.bin','ColdTangentPredictors.json','EnvelopeOptions.bin','NativeTolerance.bin',
 'NativeSolver.json','ArcCallAllowance.bin','OperationalArcRatio.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin']
raw={}
for n,r in roots.items():
    files=common+['Input/'+name for name in inputs]
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
helpers=['Stage-TestColdTangentPairs.py','CandidateTestColdTangentPairs.bas.txt','Check-TestColdTangentPairs.py',
 'Align-TestEntryTrailingBlank.py','Check-TestArcMetricCounterexample.py','ArcMetricCounterexample.json',
 'Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py','Check-TestPersistentHEvidence.py',
 'Check-TestActualReverseCapture.py','Read-TestColdRawCapture.py','Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
summary=dict(results={str(n):v for n,v in results.items()},analytical_counterexample=counter,
 raw_archive=archive.name,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),
 all12_native_full_half_roots_equilibrium_and_three_station_steel_verified=True,
 all6_pairs_material_H_advances_and_restore_same_lambda_field=True,
 fixed_elastic_reference_metric_two_of_three_pass_all_joint_gates=True,largest_step_still_rejected=True,
 existing_force_material_and_steel_tolerances_not_relaxed=True,new_joint_metric_not_yet_in_working_solver=True,
 original_lambda_projection_has_analytical_smooth_peak_obstruction=True,
 failed_custom_token_alignment_not_applied=True,strict_source_audit_passed_after_one_blank_correction=True,
 constitutive_two_accepted_half_step_error_not_verified_by_this_curve_test=True,
 no_branch_or_capacity_acceptance=True,no_working_or_release_promotion=True,
 physical_capacity_certified=False,whole_goal_completed=False,goal_active=True)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
 pause_request_revoked=True,pause_after_current_stage=False,latest_cold_tangent_metric_pairs_evidence=out.relative_to(repo).as_posix(),
 latest_native_tangent_pair_proof=results[866],latest_analytical_arc_metric_counterexample=counter,
 primary_arc_metric_review='literature/Coupled6D/NumericalControl/ArcCurveMetricReview20261126.md',
 previous_goal_turn_classification='progress:865sixunacceptedstrictrootsnowfoundbyoutgoingtangentnegativeDL (4itersperpair),allH/force/gauge/steel/restorationPASSbutoldprojectioncurvaturefails.866fixedreferencealpha=.00010929902223202709fromexistingpreparedelasticSolve independentlyverifiedresidual1.55e-15;2of3pairsPASSalljointgates(field,joint,force,gauge,steel,H,direction,locality,resolution,0iterrestore),largestcorrectlyrejects.12analyticalparabolacrossingcounterexamplesproveoldlambdaError/absDeltaλblocksanysmoothpeakcrossing;fixedscalecritical(localmedianwoulddiverge). PrimaryAbaqusRiksrefsupportscombinedspace/initiallinearreference/outgoingplasticbranch. No physicalcapacity or productionpromotion;nextactualproduction-numericalprivate867integrationandcoldverification.')
encoded=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(encoded)
print('STRICT TANGENT ROOTS/FIXED METRIC ARCHIVED',len(raw),'FILES',archive.stat().st_size,'BYTES;NO PHYSICAL LIMIT;GOAL ACTIVE')
