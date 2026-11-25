"""ДЛЯ ТЕСТОВ. Сохраняет проверенные отрицательные варианты корректора и локализацию несогласованных ветвей без объявления физического предела."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from datetime import datetime
import json,tarfile
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={870:base/'Cold870RefinedArcCorrector20_20261126',871:base/'Cold871UnchangedCorrectorControl20_20261126',872:base/'Cold872PointDerivativeAudit20_20261126',873:base/'Cold873BranchConsistentArcTangent20_20261126'}
positive=json.loads((roots[870]/'RefinedCorrectorProof.json').read_bytes());control=json.loads((roots[871]/'RefinedCorrectorProof.json').read_bytes())
point=json.loads((roots[872]/'PointDerivativeAuditProof.json').read_bytes());two=json.loads((roots[872]/'TwoBranchDirectionProof.json').read_bytes());branch=json.loads((roots[873]/'BranchConsistentTangentProof.json').read_bytes())
for p in (positive,control):
    assert p['native']['failed']==0 and p['native']['own_process_closed'] and p['new_accepted_equilibria']==0 and p['total_calls']==24
assert point['native']['passed']==2 and point['native']['failed']==0 and point['native']['own_process_closed']
assert branch['native']['failed']==0 and branch['native']['own_process_closed']
out=repo/'docs/reports/evidence/Coupled6DBranchConsistencyStudy_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
fixed=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json','CompiledSourceReuse.json','CompiledSourceReuseProof.json',
 'ParentCodeUpdate.json','ParentCompileClosure.json','docs/regression/CodeUpdate.json','Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1','Run-TestRealColdReplay.ps1','Run.ndjson',
 'RefinedCorrectorProof.json','PointDerivativeAuditProof.json','TwoBranchDirectionProof.json','BranchConsistentTangentProof.json','SourcePointDerivativeAuditProof.json','SourceVerified20ContinuityProof.json',
 'Stage-TestRefinedArcCorrector.py','Stage-TestPointDerivativeAudit.py','Align-TestPointDerivativeEntry.py','PreparedBeforePointCommentAlignment.bas','Stage-TestBranchConsistentArcTangent.py','Fix-TestBranchTangentIdentifier.py',
 'FailedCompileReservedIdentifier/CompilePoint.log','FailedCompileReservedIdentifier/CompilePointClosure.json','FailedCompileReservedIdentifier/CompilePointFailure.json',
 'tools/build_workbook/Run-Audit02ModeTests.ps1','tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
inputs=['Reverse1.bin','EnvelopeOptions.bin','NativeTolerance.bin','NativeSolver.json','ArcCallAllowance.bin','OperationalArcRatio.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin','PointDerivative.bin']
raw={}
for number,root in roots.items():
    names=fixed+['Input/'+name for name in inputs]+[p.relative_to(root).as_posix() for directory in ('src','tests','Native') for p in (root/directory).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(names)):
        if (root/name).is_file():raw[f'{number}/{name}']=(root/name).read_bytes()
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=6) as z:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(pins,indent=2,sort_keys=True)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);z.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as z:
    for name,digest in pins.items():assert sha256(z.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=['Stage-TestRefinedArcCorrector.py','Check-TestRefinedArcCorrector.py','Stage-TestPointDerivativeAudit.py','Align-TestPointDerivativeEntry.py','Check-TestPointDerivativeAudit.py','Check-TestTwoBranchDirections.py',
 'Stage-TestBranchConsistentArcTangent.py','Fix-TestBranchTangentIdentifier.py','Check-TestBranchConsistentArcTangent.py','Check-TestFixedMetricKernel.py',
 'Stage-TestEnergyNumerics.py','Check-TestPersistentHEvidence.py','Check-TestActualReverseCapture.py','Read-TestColdRawCapture.py','Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
point_summary={k:v for k,v in point.items() if k!='local_comparisons'}
summary=dict(refined_corrector=positive,unchanged_control=control,point_derivative_audit=point_summary,two_branch_direction_approximation=two,branch_consistent_predictor=branch,
    raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),raw_archive=archive.name,
    physical_materials_and_force_tolerances_unchanged=True,refined_only_corrector_not_adopted=True,
    loading_unloading_mismatch_point6_station_plus1_localized=True,self_consistent_tangent_guard_not_a_physical_failure_criterion=True,
    remaining_relative_lambda_half_gates_in_Newton=['ConfirmOriginalReturnHalf','TryEnvelopeStep'],
    separate_repeated_solution_lambda_gate_to_audit_not_mechanically_replace='TryCompletePrincipalTransition',
    prior_cold868_twenty_states_independent_stress_replay_still_required=True,
    no_release_or_main_promotion=True,physical_capacity_certified=False,whole_goal_completed=False,goal_active=True)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()};(out/'Files.json').write_text(json.dumps(top,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    latest_branch_consistency_evidence=out.relative_to(repo).as_posix(),latest_refined_corrector_proof=positive,latest_unchanged_corrector_control_proof=control,
    latest_point_derivative_audit=point_summary,latest_two_branch_direction_approximation=two,latest_branch_consistent_predictor_proof=branch,
    previous_goal_turn_classification='progress:870one-call refined primary corrector and871immutablecontrol bothactualtrue20 zero roots24calls,unhelpfulchangeNOTadopted.872readonlynativeall129points2locations bothderivatives plusdirectstressdifferences;probe refinedmatrix independentlyannihilatesdirection2.93e-17 butincomingprobe loads6/+1 whilecomputedforwarddirection unloads;twobranch independentapprox confirmsopposite signs.873actualgeneralNewton/Arc branchsignatureiteration4budgetedbuilds compiled2imports254others verified;actualnativeoutcomearchived,notphysicalfailure. Furtherduplicate relativeDeltaLambda half-gates identified in Newton H/event consumers;next reconcilecurvecontrols and investigate admissible H-directed continuation,noforcerelaxation.')
data=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('BRANCH CONSISTENCY STUDY ARCHIVED',len(raw),'FILES',archive.stat().st_size,'BYTES;NO PHYSICAL LIMIT;GOAL ACTIVE')
