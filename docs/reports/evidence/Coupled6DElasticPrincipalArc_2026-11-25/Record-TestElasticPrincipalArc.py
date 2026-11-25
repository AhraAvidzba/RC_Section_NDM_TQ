"""ДЛЯ ТЕСТОВ. Архивирует исправление спектрального перехода, независимые равновесия и регрессии без объявления физического предела."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
from io import BytesIO
import tarfile
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={836:base/'Cold836ElasticPrincipalEvent_20261125',837:base/'Cold837ElasticPrincipalEvent_20261125',838:base/'Cold838ElasticPrincipalArc_20261125'}
point=json.loads((roots[837]/'ElasticPrincipalEventProof.json').read_bytes())
walk=json.loads((roots[838]/'AutomaticLoadingWalkProof.json').read_bytes())
events=json.loads((roots[838]/'ElasticPrincipalArcEventsProof.json').read_bytes())
scope=json.loads((roots[838]/'ElasticPrincipalScopeProof.json').read_bytes())
assert point['solved_events']==6 and point['analytic_gradients_verified']==2322
assert walk['new_accepted_equilibria']==27 and len(walk['all_equilibria']['states'])==113 and walk['total_calls']==95
assert len(events['events'])==7 and events['all_actual_stress_zeros_verified']
assert not any(p['physical_capacity_certified'] for p in (point,walk,events,scope))
check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'];native={}
for n,label,passed in [(837,'RealColdReplay',1),(838,'RealColdReplay',1),(838,'OperationalArcRegressions',164)]:
    value,_=check(roots[n],label);assert value['passed']==passed and value['failed']==0 and value['own_process_closed'] and not value['timeout600']
    native[f'{n}_{label}']=value
for n in roots:
    assert json.loads((roots[n]/'CompilePointClosure.json').read_bytes())['State']=='Absent'
out=repo/'docs/reports/evidence/Coupled6DElasticPrincipalArc_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
common=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompilePointFailure.json',
    'Compile-TestExternalPoint.ps1','ParentCodeUpdate.json','ParentCompileClosure.json',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls',
    'src/Materials/CCoupledConcreteReturn.cls','src/Solver/CCoupledMaterialConstraint.cls',
    'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls',
    'src/Solver/CCoupledDeformationContinuation.cls','tests/modTestCoupledElasticState.bas',
    'tools/build_workbook/Update-IsolatedOptimizerCode.ps1','Guard-IsolatedOffice.ps1',
    'Run-TestRealColdReplay.ps1','Stage-TestElasticPrincipalEvent.py','Stage-TestElasticPrincipalArc.py',
    'Recover-TestSavedCloseCompilation.ps1','SavedCompilationRecovery.json','CodeUpdateBeforeRecovery.json']
compiled=['docs/regression/CodeUpdate.json','PointTraceSourceComposition.json','CompiledSourceReuse.json',
    'CompiledSourceReuseProof.json','CompileClosure.json']
inputs=['Reverse1.bin','ArcScanFields.bin','ArcScanCases.json','ElasticPrincipalTargets.bin','EnvelopeOptions.bin',
    'NativeTolerance.bin','NativeSolver.json','ArcCallAllowance.bin','OperationalArcRatio.bin',
    'SourceColdRawHistory.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin']
raw={}
for n,r in roots.items():
    files=common+(['Input/'+name for name in inputs])
    if n!=836:
        files+=compiled
        files += ['ElasticPrincipalEventProof.json'] if n==837 else ['AutomaticLoadingWalkProof.json','ElasticPrincipalArcEventsProof.json','ElasticPrincipalScopeProof.json']
        files += [p.relative_to(r).as_posix() for p in (r/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        p=r/name
        if p.is_file():raw[f'{n}/{name}']=p.read_bytes()
rawpins={name:sha256(data).hexdigest().upper() for name,data in raw.items()}
archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=6) as z:
    content=dict(raw);content['RawFiles.json']=(json.dumps(rawpins,indent=2,sort_keys=True)+'\n').encode('utf-8')
    for name,data in content.items():
        info=tarfile.TarInfo(name);info.size=len(data);z.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as z:
    for name,digest in rawpins.items():assert sha256(z.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=['Stage-TestElasticPrincipalEvent.py','Stage-TestElasticPrincipalArc.py','Stage-TestSavedCloseRecovery.py',
    'Recover-TestSavedCompilation.ps1','Check-TestElasticPrincipalEvent.py','Check-TestElasticPrincipalArcEvents.py',
    'Check-TestElasticPrincipalScope.py','Check-TestAutomaticLoadingWalk.py','Finalize-TestOperationalArc.py',
    'Finalize-TestStockYield.py','Inspect-TestResolutionVBA.py','Audit-TestPrivateCoupledBook.py',
    'Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py','Read-TestColdRawCapture.py',
    'Check-TestActualReverseCapture.py','Check-TestPersistentHEvidence.py','Check-TestHardHalfSteps.py',
    'Validate-TestHardPrepared.py','Run-TestOperationalArcRegressions.ps1',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
release={name:sha256((repo/name).read_bytes()).hexdigest().upper() for name in
    ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary=dict(point=point,walk=walk,spectral_events=events,source_scope=scope,native=native,release_hashes=release,
    raw_archive='RawEvidence.tar.xz',raw_files=len(rawpins),raw_uncompressed_bytes=sum(map(len,raw.values())),
    failed836_declaration_preserved=True,separate837_COM_close_failure_recovered_after_save_and_actual_source_audit=True,
    original837_compiler_invocation_not_relabelled_success=True,
    no_838_cold_capacity_search=True,not_clean_performance_benchmark=True,
    physical_capacity_certified=False,full_acceptance_completed=False,whole_goal_completed=False,goal_active=True)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
state_path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(state_path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
    pause_request_revoked=True,pause_after_current_stage=False,active_jobs=[],active_exec_session=None,all_started_stages_terminal=True,
    elastic_principal_arc_evidence=out.relative_to(repo).as_posix(),latest_verified_extended_branch=str(roots[838]),
    latest_verified_spectral_event_source=str(roots[837]),
    elastic_principal_arc_summary=dict(initial_states=86,new_equilibria=27,total_states=113,calls=95,
        last_lambda=walk['all_equilibria']['states'][-1]['lambda'],spectral_events=7,
        maximum_three_station_steel_utilization=walk['maximum_three_station_steel_utilization'],regression_passed=164,
        physical_capacity_certified=False),
    previous_goal_turn_classification='progress:835 actual49-state cold stress replay completed after explicit pause revocation. Current836-838 fixes missing generic CDP elastic principal sign event:837 solves6 frozen segments,2322 analytic gradients independent;838 passes original86 stopping point and adds27 verified equilibria to113 total,95 calls,lambda0.208818271363159;all accepted/half roots/six forces/three-station steel and7 actual stress zero events verified,164 directed regressions PASS. No material law/acceptance math changed; no physical ultimate or full acceptance.',
    next_action='Continue full original active goal; pause explicitly revoked.838 generic elastic spectral transitions verified at points3/8/31,principals1/3,both sign directions. Current branch113 ends atlambda0.208818271363159 from96-call budget(95 used),not physical failure. Reconstruct exact113 prefix from838 packets and continue unchanged integrated kernel to understand admissible postcrack path/possible later steel carrying branch and physical limit ONQ012. Also integrate spectral kernel with831/832 gated/proactive cold controller and address cold96-budget distribution honestly; authoritative cold832 still49 states,max1.1058869810286156,NumFail;835 independent replay passed. Need cold/peak/step/mesh/external validation and complete original LoadPaths/hoop/UI/CAD/cracks/stability/performance/workbook/full scoped acceptance. Production/main/release unchanged.836 compile syntax failure preserved;837 saved compilation had separateRPC Close failure,own process closed,recovered print areas/passport,actual256 sources verified;838 normal compile/save/closure verified. GitHub92948eb/9b252c9 not yet sent:HTTPS HTTP408 twice,SSH official host verified but no publickey authentication; retry latest checkpoint using working authorized HTTPS,do not claim remote updated without verification.')
state['terminal_sessions_this_stage']=sorted(set(state.get('terminal_sessions_this_stage',[])+[7564,68396,90800,29834,35024,62959,50587,86488]))
encoded=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (state_path,here/'Continuation666_Resume20261123.json'):target.write_bytes(encoded)
print('ARCHIVED',len(rawpins),'RAW FILES',sum(map(len,raw.values())),'BYTES IN',archive.stat().st_size,'BYTE TAR.XZ;',len(pins),'TOP PINS;GOAL ACTIVE')
