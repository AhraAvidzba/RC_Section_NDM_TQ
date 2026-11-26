"""ДЛЯ ТЕСТОВ. Сохраняет рабочие дуговые диагностики, отдельное продолжение и проверенный холодный результат без объявления физического предела."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import ctypes,json,tarfile

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={893:base/'Cold893Arc56Diagnostics_20261126',894:base/'Cold894Cold56ArcWalk_20261126',
       895:base/'Cold895FreshSpatialArc_20261126',896:base/'Cold896FreshArcRegressions_20261126',
       897:base/'Cold897FreshArcCold96_20261126',898:base/'Cold898FreshArcColdStress_20261126'}
diag=json.loads((roots[893]/'Arc56DiagnosticsProof.json').read_bytes())
walk=json.loads((roots[894]/'Cold56ArcWalkProof.json').read_bytes())
selector=json.loads((roots[895]/'FreshSpatialArcProof.json').read_bytes())
cold=json.loads((roots[897]/'FreshArcCold96Proof.json').read_bytes())
stress=json.loads((roots[898]/'VerifiedColdStressReplayProof.json').read_bytes())
assert len(diag['cases'])==8 and all(c['solved'] for c in diag['cases'])
assert walk['new_accepted_equilibria']==14 and walk['total_calls']==94
assert walk['independently_rebuilt_half_root_count']==14
assert selector['native']['passed']==23 and selector['old_exact_retained_arc_gate_unchanged']
reg,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](roots[896],'OperationalArcRegressions')
assert reg==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
assert cold['accepted_states']==stress['source_states'] and stress['native']['own_process_closed']
assert stress['all_three_station_steel_limits_checked']
statepath=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json'
state=json.loads(statepath.read_bytes());assert not state['active_jobs']

class TestMemory(ctypes.Structure):
    """Проверяет физический резерв и запас commit перед сжатием двоичных VBA-проектов."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory)
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>3*1024**3 and memory.free_page>3*1024**3
out=repo/'docs/reports/evidence/Coupled6DFreshArcCold_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
fixed=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
 'CompiledSourceReuse.json','ParentCodeUpdate.json','ParentCompileClosure.json','docs/regression/CodeUpdate.json',
 'Arc56DiagnosticsProof.json','Cold56ArcWalkProof.json','FreshSpatialArcProof.json','FreshArcCold96Proof.json','FreshArcColdActualComparison.json',
 'VerifiedColdStressReplayProof.json','ReplaySourceScope.json','SourceColdProof.json','PriorDirectedRegressionProof.json',
 'Run-TestHMixedStock.ps1','Run-TestRealColdReplay.ps1','Run-TestOperationalArcRegressions.ps1',
 'Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1','tools/build_workbook/Update-IsolatedOptimizerCode.ps1',
 'tools/build_workbook/Run-Audit02ModeTests.ps1']
raw={}
for number,root in roots.items():
    names=fixed+[p.relative_to(root).as_posix() for directory in ('src','tests','Native')
                 for p in (root/directory).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    names+=['Input/'+n for n in ('Reverse1.bin','SourceColdRawHistory.bin','EnvelopeOptions.bin','NativeTolerance.bin','NativeSolver.json')]
    for name in sorted(set(names)):
        if (root/name).is_file():raw[f'{number}/{name}']=(root/name).read_bytes()
    if number in (893,894,895,898):
        with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:
            raw[f'{number}/ActualCompiledVBA.bin']=book.read('xl/vbaProject.bin')
if (roots[897]/'SavedWorkbook.xlsm').is_file():
    with ZipFile(roots[897]/'SavedWorkbook.xlsm') as book:raw['897/ActualSavedVBA.bin']=book.read('xl/vbaProject.bin')
for dirname,filename in [('Cold835GatedColdReplay_20261125','src/Solver/CCoupledElasticState.cls'),
                         ('Cold835GatedColdReplay_20261125','Native/EnvelopeEnergy.Accepted.Operators.csv'),
                         ('Cold880CriticalHCold96_20261126','Native/Stock.RawHistory.bin'),
                         ('Cold890StrictGrowthCold96_20261126','Native/Stock.RawHistory.bin'),
                         ('Cold888StrictEnvelopeGrowth_20261126','src/Solver/CCoupledDeformationContinuation.cls')]:
    raw[f'Referenced/{dirname}/{filename}']=(base/dirname/filename).read_bytes()
pins={n:sha256(b).hexdigest().upper() for n,b in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=9) as bundle:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);bundle.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as bundle:
    for name,digest in pins.items():assert sha256(bundle.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=['Stage-TestArc56Diagnostics.py','Align-TestArc56Comments.py','Check-TestArc56Diagnostics.py',
 'Stage-TestCold56ArcWalk.py','Check-TestCold56ArcWalk.py','Check-TestFixedMetricKernel.py',
 'Stage-TestFreshSpatialArc.py','Align-TestFreshSpatialArcComments.py','Check-TestFreshSpatialArc.py',
 'Stage-TestPreparedRegressionRun.py','Stage-TestFreshArcCold96.py','Check-TestFreshArcCold96.py',
 'Stage-TestStrictGrowthCold96.py','Check-TestStrictGrowthCold96.py','Check-TestCriticalHCold96.py',
 'Check-TestFixedMetricCold.py','Check-TestSpatialProjectionCompatibility.py',
 'Stage-TestFreshArcColdStress.py','Check-TestFreshArcColdStress.py','Stage-TestStrictGrowthColdStress.py',
 'Check-TestStrictGrowthColdStress.py','Stage-TestImmutableColdStressReplay.py','Check-TestCriticalHColdStress.py',
 'Check-TestVerifiedColdStressReplay.py','Check-TestSharedCurveHControl.py','Check-TestPersistentHEvidence.py',
 'Read-TestColdRawCapture.py','Stage-TestColdRawReplay.py','Check-TestActualReverseCapture.py',
 'Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py','Stage-TestEnergyNumerics.py',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
summary=dict(arc56_diagnostics893=diag,separate_follow_on894=walk,fresh_selector895=selector,
 regressions896=reg,cold897=cold,cold_stress898=stress,raw_archive=archive.name,raw_files=len(raw),
 raw_bytes=sum(map(len,raw.values())),xz_preset=9,free_physical_before=memory.free_physical,
 free_commit_before=memory.free_page,full_acceptance_not_run=True,physical_capacity_certified=False,
 no_release_or_main_promotion=True,whole_goal_complete=False)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
 latest_fresh_arc_evidence=out.relative_to(repo).as_posix(),latest_fresh_arc_cold897=cold,
 latest_fresh_arc_stress898=stress,latest_fresh_selector895=selector,
 latest_cold56_arc_follow_on894=walk,previous_goal_turn_classification='progress: actual arc56 roots and14 follow-on pairs independently equilibrated; fresh spatial context compiled and164 regressions passed; cold897 and complete frozen stress replay898 verified. No confirmed physical ultimate.',
 next_action='Audit actual outgoing Jacobian/residual at cold897 state32, especially row53. Fresh policy activated3times but no improvement:32states/max1.1042833676987807/820.4443359375s/NumFail versus56states/731.138671875s in890; no promotion. Separate follow-on894 is not capacity. Preserve exact material/history and96/30 budgets. Full SOW remains active; pure bending CDP versus SP deferred.')
data=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for file in (statepath,here/'Continuation666_Resume20261123.json'):file.write_bytes(data)
print('ARCHIVED',len(raw),'RAW FILES',archive.stat().st_size,'BYTES;GOAL ACTIVE;NO PHYSICAL CAPACITY CLAIM')
