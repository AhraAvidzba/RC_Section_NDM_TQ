"""ДЛЯ ТЕСТОВ. Архивирует холодный исход и проверку всех25 состояний, включая точную сохранённую книгу; физический предел не объявляет."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import ctypes,json,tarfile,numpy as np
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
source=base/'Cold920AffineRecoveryCold96_20261126';replay=base/'Cold921AffineRecoveryColdStress_20261126'
cold=json.loads((source/'AffineRecoveryCold96Proof.json').read_bytes());stress=json.loads((replay/'VerifiedColdStressReplayProof.json').read_bytes())
assert cold['cold']['all_saved_sources_exact']==256 and cold['cold']['native']['own_process_closed']
assert cold['cold']['capacity']=='NumFail' and cold['accepted_states']==25
assert stress['source_states']==25 and stress['all_raw_fields_fractions_and_lambdas_exact']
assert stress['all_six_forces_and_all_56_physical_and4_gauge_equations'] and stress['all_three_station_steel_limits_checked']
assert stress['native']['own_process_closed'] and stress['material_and_kinematic_files_and_replay_methods_exact_source']
read=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory']
frames,_=read(source/'Native/Stock.RawHistory.bin');current=frames[-1]
prior=base/'Cold904ConsistentRefinementCold96_20261126';old=read(prior/'Native/Stock.RawHistory.bin')[0][-1]
prefix=0
for i in range(min(len(current['fields']),len(old['fields']))):
    if not all(np.array_equal(current[k][i],old[k][i]) for k in ('fields','fractions','lambdas')):break
    prefix+=1
assert prefix==7
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](replay,'EnvelopeEnergy.Accepted')
assert len(fields)==25
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](replay/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)]);physical=[]
for i in (int(np.argmax(current['lambdas'])),24):
    strain=np.einsum('spcd,d->spc',op,current['fields'][i]);tensor=np.zeros((3,25,3,3))
    for c in range(3):tensor[:,:,c,c]=strain[:,:25,c]
    for c,(a,b) in enumerate(((0,1),(0,2),(1,2)),3):tensor[:,:,a,b]=tensor[:,:,b,a]=strain[:,:25,c]/2
    principal=np.linalg.eigvalsh(tensor);steel=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
    physical.append(dict(state=i+1,lambda_value=float(current['lambdas'][i]),concrete_principal_min=float(principal.min()),
        concrete_principal_max=float(principal.max()),max_abs_long_steel_strain=float(np.max(abs(steel[:,long]))),
        max_abs_hoop_strain=float(np.max(abs(steel[:,~long])))))
lines=(source/'Native/Stock.Checkpoints.csv').read_text('cp1251').splitlines();details=[x for x in lines if x.startswith('DETAILS;')]
assert details and 'уравнение 55' in details[-1] and 'продольный градиент My' in details[-1]
statepath=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statepath.read_bytes());assert not state['active_jobs']
class TestMemory(ctypes.Structure):
    """Проверяет RAM/commit перед большим словарём XZ и копированием точной контрольной книги."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory);assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>3*1024**3 and memory.free_page>3*1024**3
out=repo/'docs/reports/evidence/Coupled6DAffineRecoveryCold_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
raw={};fixed=['Preparation.json','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
 'docs/regression/CodeUpdate.json','CompiledSourceReuse.json','ParentCodeUpdate.json','ParentCompileClosure.json',
 'Input/Reverse1.bin','Input/SourceColdRawHistory.bin','SourceColdProof.json','PriorDirectedRegressionProof.json',
 'AffineRecoveryCold96Proof.json','VerifiedColdStressReplayProof.json','ReplaySourceScope.json',
 'Run-TestHMixedStock.ps1','Run-TestRealColdReplay.ps1','Guard-IsolatedOffice.ps1',
 'tools/build_workbook/Run-Audit02ModeTests.ps1','tools/build_workbook/Run-FocusedKaniWatchdog.ps1']
for number,root in ((920,source),(921,replay)):
    files=fixed+[p.relative_to(root).as_posix() for folder in ('src','tests','Native') for p in (root/folder).rglob('*') if p.is_file()]
    for name in sorted(set(files)):
        if (root/name).is_file():raw[f'{number}/{name}']=(root/name).read_bytes()
    with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:raw[f'{number}/ActualCompiledVBA.bin']=book.read('xl/vbaProject.bin')
saved=(source/'SavedWorkbook.xlsm').read_bytes();assert sha256(saved).hexdigest().upper()==cold['cold']['saved_book_sha256']
raw['920/ActualSavedWorkbook.xlsm']=saved
with ZipFile(BytesIO(saved)) as book:raw['920/ActualSavedVBA.bin']=book.read('xl/vbaProject.bin')
raw['920/LastFailureDetails.txt']=(details[-1]+'\n').encode('utf-8')
oldproof=json.loads((prior/'ConsistentRefinementCold96Proof.json').read_bytes())
historical=read(base/'Cold853LoadingSecantCold_20261125/Native/Stock.RawHistory.bin')[0][-1]
comparison=dict(exact_prefix_states_with904=prefix,prior904_accepted_states=len(old['fields']),prior904_maximum_lambda=float(max(old['lambdas'])),
    prior904_seconds=oldproof['cold']['seconds'],maximum_lambda_change_from904_percent=100*(max(current['lambdas'])/max(old['lambdas'])-1),
    historical853_maximum_lambda=float(max(historical['lambdas'])),historical853_is_different_algorithm_not_clean_comparison=True,
    clean_speed_comparison_not_claimed=True,cold_overlapped_archiving_completed914_to919=True)
pins={n:sha256(b).hexdigest().upper() for n,b in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=9) as bundle:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);bundle.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as bundle:
    for name,digest in pins.items():assert sha256(bundle.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=['Stage-TestAffineRecoveryCold96.py','Check-TestAffineRecoveryCold96.py','Stage-TestStrictGrowthCold96.py',
 'Check-TestStrictGrowthCold96.py','Check-TestCriticalHCold96.py','Check-TestFixedMetricCold.py',
 'Check-TestSpatialProjectionCompatibility.py','Stage-TestAffineRecoveryColdStress.py','Check-TestAffineRecoveryColdStress.py',
 'Stage-TestStrictGrowthColdStress.py','Check-TestStrictGrowthColdStress.py','Stage-TestImmutableColdStressReplay.py',
 'Check-TestCriticalHColdStress.py','Check-TestVerifiedColdStressReplay.py','Read-TestColdRawCapture.py',
 'Check-TestActualReverseCapture.py','Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py','Check-TestPersistentHEvidence.py',
 'Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py','Inspect-TestResolutionVBA.py','Canonical-TestCoreVBA.py',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
summary=dict(cold=cold,verified_stress_replay=stress,independently_reassembled_equilibrium=mechanics,physical_selected_states=physical,
 comparison=comparison,entire_new_cold_trajectory_stress_replay_completed=True,
 raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),raw_archive=archive.name,raw_archive_bytes=archive.stat().st_size,
 all_raw_hashes_read_back=True,exact_saved_workbook_in_archive=True,free_physical_before=memory.free_physical,free_commit_before=memory.free_page,
 last_failure_details_file='920/LastFailureDetails.txt',failure_details_may_be_cached_actual_context_capture_still_required=True,
 final_residual_is_large_not_at_old_point5_machine_precision=True,
 all_started_stages_terminal=True,no_working_release_main_promotion=True,physical_capacity_certified=False,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),latest_cold920_evidence=out.relative_to(repo).as_posix(),
 latest_completed_cold920=cold,latest_verified_stress_replay921=stress,
 latest_cold920_material_replay_completed=True,all_started_stages_terminal=True,
 previous_goal_turn_classification='progress:914 reproduced actual failed matrix/residual and found83%descent;915/918 restored previously unavailableCase7Full, all5 Arc acceptances retained and917/919 each164PASS.916 five matched-branch affine predictions independently equilibrated.920 true cold source918 completed787.162109375s/directOK/capacityNumFail/25states/max1.1052194649105769; actual256 save/reopen exact.921 fresh replay independently verified every new state/all6/56+4/three-stationsteel. First7states exact904, then trajectory differs; final diagnostic row55/point11-station-1 is a different large-residual failure, exact failed context still required. No physical ultimate or release promotion; all processes closed.')
data=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (statepath,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('ARCHIVED COLD920/REPLAY921',len(raw),'RAW',archive.stat().st_size,'BYTES;EXACT SAVED BOOK;ALL25 VERIFIED;ULTIMATE STILL UNCONFIRMED')
