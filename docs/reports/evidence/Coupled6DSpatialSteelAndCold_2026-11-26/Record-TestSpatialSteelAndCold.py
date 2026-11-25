"""ДЛЯ ТЕСТОВ. Архивирует завершённый холодный расчёт, повтор напряжений и контроль стального предела трёх станций без объявления полного выполнения ТЗ."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import json,tarfile,re
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={880:base/'Cold880CriticalHCold96_20261126',881:base/'Cold881SpatialSteelLimits_20261126',
       882:base/'Cold882CriticalHColdStress_20261126',883:base/'Cold883SpatialSteelRegressions_20261126',884:base/'Cold884SpatialSteelBoundary_20261126'}
native=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
steel,_=native(roots[881],'SpatialSteel');boundary,_=native(roots[884],'SpatialBoundary');reg,_=native(roots[883],'OperationalArcRegressions')
assert steel==dict(passed=31,failed=0,macros=1,timeout600=False,own_process_closed=True)
assert boundary==dict(passed=12,failed=0,macros=1,timeout600=False,own_process_closed=True)
assert reg==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
cold=json.loads((roots[880]/'CriticalHCold96Proof.json').read_bytes())
stress=json.loads((roots[882]/'VerifiedColdStressReplayProof.json').read_bytes())
assert cold['accepted_states']==stress['source_states']==56
assert cold['cold']['capacity']=='NumFail' and cold['cold']['all_saved_sources_exact']==256
assert cold['cold']['native']['own_process_closed'] and stress['native']['own_process_closed']
assert stress['all_three_station_steel_limits_checked'] and stress['all_raw_fields_fractions_and_lambdas_exact']
assert stress['material_and_kinematic_files_and_replay_methods_exact_source']
audit=json.loads((roots[881]/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and len(audit['imported_this_pass'])==6
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==250
audit2=json.loads((roots[884]/'PointTraceSourceComposition.json').read_bytes())
assert audit2['components']==256 and audit2['imported_this_pass']==['modTestCoupledElasticState']
text=(roots[884]/'Native/SpatialBoundary.txt').read_text('utf-8-sig')
brackets=[tuple(map(float,m)) for m in re.findall(r'SPATIAL_STEEL_LIMIT: sign=(-?1); found=True; low=([0-9.E+-]+); high=([0-9.E+-]+);',text)]
assert len(brackets)==2 and {a for a,b,c in brackets}=={-1,1}
assert all(0<high-low<=.0001 for sign,low,high in brackets)
out=repo/'docs/reports/evidence/Coupled6DSpatialSteelAndCold_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
fixed=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
       'CompiledSourceReuse.json','ParentCodeUpdate.json','ParentCompileClosure.json','docs/regression/CodeUpdate.json',
       'CriticalHCold96Proof.json','VerifiedColdStressReplayProof.json','ReplaySourceScope.json','SourceColdProof.json',
       'PriorDirectedRegressionProof.json','Run-TestHMixedStock.ps1','Run-TestRealColdReplay.ps1','Run-TestOperationalArcRegressions.ps1',
       'Run-TestSpatialSteelLimits.ps1','Run-TestSpatialSteelBoundary.ps1','Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1',
       'PreparedSpatialSteelBeforeCommentAlignment.bas','PreparedSpatialBoundaryBeforeCommentAlignment.bas',
       'tools/build_workbook/Update-IsolatedOptimizerCode.ps1','tools/build_workbook/Run-Audit02ModeTests.ps1']
raw={}
for number,root in roots.items():
    names=fixed+[p.relative_to(root).as_posix() for d in ('src','tests','Native','CompileFailurePrivateInstanceAccess') for p in (root/d).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    names+=['Input/'+n for n in ('Reverse1.bin','EnvelopeOptions.bin','NativeTolerance.bin','NativeSolver.json','SourceColdRawHistory.bin')]
    for name in sorted(set(names)):
        if (root/name).is_file():raw[f'{number}/{name}']=(root/name).read_bytes()
    with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:raw[f'{number}/ActualCompiledVBA.bin']=book.read('xl/vbaProject.bin')
with ZipFile(roots[880]/'SavedWorkbook.xlsm') as book:raw['880/ActualSavedVBA.bin']=book.read('xl/vbaProject.bin')
pins={n:sha256(b).hexdigest().upper() for n,b in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=6) as bundle:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);bundle.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as bundle:
    for name,digest in pins.items():assert sha256(bundle.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=['Stage-TestSpatialSteelLimits.py','Complete-TestSpatialSteelStage.py','Repair-TestSpatialSteelCompilation.py',
         'Repair-TestSpatialSteelRunner.py','Finalize-TestSpatialSteel.py','Align-TestSpatialSteelComments.py',
         'Stage-TestSpatialSteelBoundary.py','Align-TestSpatialBoundaryComments.py','Stage-TestPreparedRegressionRun.py',
         'Check-TestCriticalHCold96.py','Check-TestFixedMetricCold.py','Stage-TestImmutableColdStressReplay.py',
         'Check-TestCriticalHColdStress.py','Check-TestVerifiedColdStressReplay.py','Stage-TestEnergyNumerics.py',
         'Stage-TestColdRawReplay.py','Read-TestColdRawCapture.py','Check-TestActualReverseCapture.py',
         'Check-TestPersistentHEvidence.py','Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
summary=dict(cold880=cold,stress882=stress,spatial_steel881=steel,side_boundary884=boundary,regressions883=reg,
    side_steel_boundary_brackets=brackets,actual_spatial_steel_component_audit=audit,
    physical_HardShear_limit_certified=False,synthetic_side_limit_confirmed_not_concrete_capacity=True,
    synthetic_point_law='diagonal30000,shear15000,axial30000*epsilon+1e10*epsilon^3;not_a_concrete_calibration',
    raw_archive=archive.name,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),
    full_acceptance_not_run=True,no_release_or_main_promotion=True,goal_active=True,
    remaining=['HardShear confirmed ultimate;larger local state count did not improve maximum lambda or time',
      'independent raw-field/stress replay of new synthetic side-limit test;its current comment incorrectly says linear',
      'side metadata in load-controlled steel-limit diagnostics and detailed output;native end-to-end reports for side event',
      'step integration and two sequential accepted half-steps;mesh and step sensitivity;ONQ010/012',
      'external material-matched controls;full SOW and scoped acceptance;target10seconds;main/release promotion'])
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes());assert not state['active_jobs']
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    active_cold_exec_session=None,active_cold_source_root=None,latest_completed_cold880=cold,latest_completed_cold880_stress=stress,
    latest_spatial_steel_evidence=out.relative_to(repo).as_posix(),latest_spatial_steel_kernel_root=str(roots[881]),
    latest_spatial_steel_native=steel,latest_side_steel_boundary_native=boundary,latest_spatial_steel_regressions=reg,
    generic_runtime_three_station_steel_limit_audit_open=['independent raw replay and detailed reporting for side-governing limit;core aggregate and native localization implemented'],
    previous_goal_turn_classification='progress:cold880 terminal56states/max1.104283/748s/directOK/capacityNumFail; savedreopened256verified; pure frozenstress882 verifies all56sixforces/56+4/all3steel. New881 aggregates all3station steel in common response, station projection/direct boundaryscreen/deformation event side/transverse fields, native31PASS;883unchanged164PASS. Native884real3stationQx± force solver localizes side-steel physical boundary [1.63922119140625,1.639312744140625],12PASS. Failed compile privateaccess and inheritedrunner preserved/fixed. MainHardShear limit still unconfirmed;no release/main promotion.',
    next_action='Capture raw fields/operators and independently verify the new synthetic side-governing steel boundary; correct its test-law description, add side metadata in load-controlled event reports, then return to cold continuation bottleneck. Do not repeat long cold run without a specific numerical hypothesis.')
data=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for file in (path,here/'Continuation666_Resume20261123.json'):file.write_bytes(data)
print('COLD880 VERIFIED;STEEL31+12;REGRESSION164;ARCHIVED',len(raw),'FILES',archive.stat().st_size,'BYTES;GOAL ACTIVE;NO RELEASE PROMOTION')
