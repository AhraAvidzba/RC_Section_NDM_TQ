"""ДЛЯ ТЕСТОВ. Закрепляет разбор разгрузочной ветви и проверенный возврат к нагружению через существующие материальные уравнения."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
from io import BytesIO
import json,tarfile

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={839:base/'Cold839ContinuedSpectralWalk_20261125',840:base/'Cold840EnvelopeReentry_20261125',
    841:base/'Cold841HGuideMaterialEvent_20261125',842:base/'Cold842HReentryWalk_20261125'}
reentry=json.loads((roots[840]/'EnvelopeReentryProof.json').read_bytes())
face=json.loads((roots[841]/'HGuideMaterialEventProof.json').read_bytes())
walk=json.loads((roots[842]/'HReentryWalkProof.json').read_bytes())
assert reentry['last_material_growth_state']==70 and reentry['all8_separate_equilibria_independently_verified']
assert all(not p['verified'] for p in reentry['strict_full_half_pairs'])
assert face['verified_local_bridge_to_loading_boundary'] and face['zero_iteration_exact_restore_verified']
assert face['event_reaches_new_plastic_loading_with_growing_H']
assert walk['all_new_states_have_strictly_growing_H'] and walk['walk']['new_accepted_equilibria']==8
assert walk['walk']['total_calls']==27 and walk['first_loading_departure_kind']==4
assert all(not p['physical_capacity_certified'] for p in (reentry,face,walk))
assert not any((roots[839]/'Native').iterdir())
native={};check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
for n in (840,841,842):
    value,_=check(roots[n],'RealColdReplay')
    assert value==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
    native[str(n)]=value
for n,root in roots.items():
    assert json.loads((root/'CompilePointClosure.json').read_bytes())['State']=='Absent'
    audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
    assert audit['components']==256
    assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==256-len(audit['imported_this_pass'])
out=repo/'docs/reports/evidence/Coupled6DHReentry_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
common=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
    'CompiledSourceReuse.json','CompiledSourceReuseProof.json','ParentCodeUpdate.json','ParentCompileClosure.json',
    'docs/regression/CodeUpdate.json','Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1','Run-TestRealColdReplay.ps1',
    'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledDeformationContinuation.cls',
    'src/Solver/CCoupledEnvelopeEquations.cls','src/Solver/CCoupledMaterialConstraint.cls',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls','src/Materials/CCoupledConcreteReturn.cls',
    'tests/modTestCoupledElasticState.bas','tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
stage={839:'Stage-TestContinuedSpectralWalk.py',840:'Stage-TestEnvelopeReentry.py',841:'Stage-TestHGuideMaterialEvent.py',842:'Stage-TestHReentryWalk.py'}
selected={840:['EnvelopeReentryProof.json'],841:['HGuideMaterialEventProof.json'],842:['HReentryWalkProof.json','AutomaticLoadingWalkProof.json']}
inputs=['Reverse1.bin','HMaterialGuide.bin','ArcCallAllowance.bin','OperationalArcRatio.bin','EnvelopeOptions.bin',
    'NativeTolerance.bin','NativeSolver.json','SourceColdRawHistory.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin']
raw={}
for n,root in roots.items():
    names=common+[stage[n]]+selected.get(n,[])+['Input/'+name for name in inputs]
    if n!=839:names += [p.relative_to(root).as_posix() for p in (root/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(names)):
        p=root/name
        if p.is_file():raw[f'{n}/{name}']=p.read_bytes()
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()}
archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=6) as z:
    content=dict(raw);content['RawFiles.json']=(json.dumps(pins,indent=2,sort_keys=True)+'\n').encode('utf-8')
    for name,data in content.items():
        info=tarfile.TarInfo(name);info.size=len(data);z.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as z:
    for name,digest in pins.items():assert sha256(z.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=list(stage.values())+['Check-TestEnvelopeReentry.py','Check-TestHGuideMaterialEvent.py','Check-TestHReentryWalk.py',
    'TestUnacceptedMechanics.py','Check-TestAutomaticLoadingWalk.py','Stage-TestColdRawReplay.py','Stage-TestEnergyNumerics.py',
    'Read-TestColdRawCapture.py','Check-TestActualReverseCapture.py','Check-TestPersistentHEvidence.py','Check-TestHardHalfSteps.py',
    'Validate-TestHardPrepared.py','Finalize-TestOperationalArc.py','Finalize-TestStockYield.py','Inspect-TestResolutionVBA.py',
    'Audit-TestPrivateCoupledBook.py',Path(__file__).name]
for name in sorted(set(helpers)):(out/name).write_bytes((here/name).read_bytes())
release={name:sha256((repo/name).read_bytes()).hexdigest().upper() for name in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary=dict(reentry=reentry,loading_face=face,automatic_departure=walk,native=native,release_hashes=release,
    raw_archive=archive.name,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),
    stage839_compiled_and_source_audited_only=True,stage839_continuation_not_run_after_unloading_identified=True,
    original_equilibrium_states71_to113_not_reclassified_as_solver_failure=True,
    existing_controller_excludes_arcRecovered_from_IsEnvelopeUnloading_guard=True,
    automatic_H_guide_admission_not_yet_integrated_into_general_controller=True,
    original_cold_maximum_not_increased=True,not_cold_capacity_search=True,
    no_working_material_or_solver_equations_changed=True,diagnostic840_actual16byte_magic='NDMHREENTRY00000',
    physical_capacity_certified=False,full_acceptance_completed=False,whole_goal_completed=False,goal_active=True)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
    pause_request_revoked=True,pause_after_current_stage=False,active_jobs=[],active_exec_session=None,all_started_stages_terminal=True,
    h_reentry_evidence=out.relative_to(repo).as_posix(),latest_verified_active_loading_branch=str(roots[842]),
    latest_verified_loading_boundary=str(roots[841]),
    h_reentry_summary=dict(original_last_material_growth_state=70,original_frozen_H_later_states=43,
        valid_unaccepted_H_roots=8,H_half_pairs_not_admitted=3,event_lambda=face['event_lambda'],
        event_point=face['selected_point'],event_side=face['selected_side'],event_field_half_error=face['field_half_error'],
        resumed_loading_states=8,resumed_loading_total_states=79,calls=27,last_lambda=walk['walk']['all_equilibria']['states'][-1]['lambda'],
        strictly_growing_H=True,physical_capacity_certified=False),
    previous_goal_turn_classification='progress:839 exact113 prefix/source rename audited;840 independently identifies frozen energy coordinates in all129 point/stations after70 and verifies8 H roots but rejects all3 strict full/half pairs as admission evidence.841 existing material-event solver reaches new point25/-1 loading boundary atlambda'+str(face['event_lambda'])+',field half error'+str(face['field_half_error'])+',strict forces/gauges/steel and exact restore0 independently verified,H grows.842 automatically departs via existing recent-loading control,then7 ordinary arcs;8 new equilibria27 calls andallhalfroots/forces/three-station steel verified,allnewHstrictlygrows. Wrong late unloading not confused with physical capacity;working equations unchanged.',
    next_action='Continue full original active goal. Authoritative LOADING path is842: original70 + independently verified841 H-guided face +8 automatic outgoing states,total79,lastlambda'+str(walk['walk']['all_equilibria']['states'][-1]['lambda'])+'. Reconstruct exact79 prefix from842 raw packets and continue while monitoring H. Original838 states71..113 are valid declining-load equilibria with all129 energy coords frozen after70; DO NOT extend their unloading towardzero as capacity evidence.840 strict H roots exist but full/half errors.1156/.2367/.4762 fail; do not admit them directly.841 working TryLoadingYieldBoundary with strict840 H guide selects25/-1 face,verifiedpairHgrowth andzero-iterationrestore.842 recent boundary guards pass;first3 outgoing attempts fail at ratios1,.5,.25,4th.125succeeds(kind4),next7kind1growingH. Next implement automatic H-guide material-event recovery/gating in common CArc/controller: current IsEnvelopeUnloading explicitly excludes arcRecovered; include arc candidate H-progress checks and route through verified material boundary,retain original30iteration phase/96calls/field/lambda/force/gauge/steel criteria and exactrestore. Count H guide solve,event,half,restore within shared budget;do not merely bypass criteria or accept isolated H roots. Test reference source70 case and actual cold gated831/832 controller. Cold832 stillNumFail49/max1.1058869810286156;physical ONQ010/012,step/mesh/external/full original production/LoadPaths/hoop/UI/CAD/cracks/stability/performance/workbook and full scoped acceptance remain.839 compiled-only continuation deliberatelynotrun;test-only standalone entry renamed TestLocateElasticPrincipalBoundary;allnewworkingmath/materialsunmodified. Main/release unchanged. Latest successful delivery55074b2 used HTTPS HTTP/2 +postBuffer536870912 +empty Expect;retry latest checkpoint and verify actualremote ref.')
state['terminal_sessions_this_stage']=sorted(set(state.get('terminal_sessions_this_stage',[])+[49446,96349,36943,37102,52825,47167,70629,2570,61773,86463,10660]))
encoded=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(encoded)
print('H REENTRY ARCHIVED:',len(raw),'RAW FILES;',archive.stat().st_size,'BYTE TAR.XZ;ALL8 LOADING STATES VERIFIED;GOAL ACTIVE')
