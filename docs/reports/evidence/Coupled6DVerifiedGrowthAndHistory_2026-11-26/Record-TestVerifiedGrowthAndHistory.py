"""ДЛЯ ТЕСТОВ. Архивирует независимые стальные корни, H-пары, последовательное интегрирование и отрицательный холодный результат неиспользованного H-ускорения."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import ctypes,json,tarfile,re,numpy as np
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={885:base/'Cold885SpatialSteelRawControl_20261126',886:base/'Cold886Cold56LargerHPairs_20261126',
       887:base/'Cold887VerifiedEnvelopeGrowth_20261126',888:base/'Cold888StrictEnvelopeGrowth_20261126',
       889:base/'Cold889StrictGrowthRegressions_20261126',890:base/'Cold890StrictGrowthCold96_20261126',
       891:base/'Cold891SequentialAcceptedH_20261126',892:base/'Cold892StrictGrowthColdStress_20261126'}
steel=json.loads((roots[885]/'SpatialSteelRawControlProof.json').read_bytes())
pairs=json.loads((roots[886]/'Cold56LargerHPairsProof.json').read_bytes())
seq=json.loads((roots[891]/'SequentialAcceptedHProof.json').read_bytes())
cold=json.loads((roots[890]/'StrictGrowthCold96Proof.json').read_bytes())
stress=json.loads((roots[892]/'VerifiedColdStressReplayProof.json').read_bytes())
assert steel['native']['passed']==14 and len(steel['independent_zero_start_boundary_roots'])==2
assert pairs['verified_pairs']==4 and [(c['ratio'],c['solved']) for c in pairs['cases']]==[(1,True),(2,True),(4,True),(8,True),(16,False),(32,False)]
assert [(c['case'],c['solved']) for c in seq['cases']]==[(1,True),(2,True),(3,True),(4,False)]
reg,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](roots[889],'OperationalArcRegressions')
assert reg==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
assert cold['accepted_states']==stress['source_states']==56 and cold['cold']['capacity']=='NumFail' and cold['cold']['all_saved_sources_exact']==256
assert stress['native']['own_process_closed'] and stress['all_three_station_steel_limits_checked']
read=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'];old,_=read(base/'Cold880CriticalHCold96_20261126/Native/Stock.RawHistory.bin');current,_=read(roots[890]/'Native/Stock.RawHistory.bin')
for key in ('fields','fractions','lambdas'):np.testing.assert_array_equal(old[-1][key],current[-1][key])
details=(roots[890]/'Native/Stock.Checkpoints.csv').read_text('cp1251')
assert 'увеличенных H-шагов с проверенной половиной' not in details
last=[line for line in details.splitlines() if line.startswith('DETAILS;')][-1]
assert 'Мера деформации: пространственная мера проверенной знаковой дуги' in last
counts=[int(x) for x in re.findall(r'; проб: (\d+)',details)];assert counts and max(counts)==96
comparison=dict(all56_fields_fractions_lambdas_bit_identical880=True,verified_accelerated_H_attempts=0,
    active_final_control='пространственная мера проверенной знаковой дуги',global_probes_max=96,
    observed_seconds=cold['cold']['seconds'],previous_observed_seconds=748.125,
    clean_performance_comparison_not_claimed=True,own_study891_overlapped_cold=True,
    H_growth_change_did_not_affect_actual_cold_path=True,physical_capacity_not_certified=True)
class TestMemory(ctypes.Structure):
    """Проверяет резерв памяти перед сжатием фактических VBA-проектов большим словарём."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory);assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>3*1024**3 and memory.free_page>3*1024**3
out=repo/'docs/reports/evidence/Coupled6DVerifiedGrowthAndHistory_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
fixed=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
 'CompiledSourceReuse.json','ParentCodeUpdate.json','ParentCompileClosure.json','docs/regression/CodeUpdate.json',
 'SpatialSteelRawControlProof.json','Cold56LargerHPairsProof.json','SequentialAcceptedHProof.json','StrictGrowthCold96Proof.json',
 'VerifiedColdStressReplayProof.json','ReplaySourceScope.json','SourceColdProof.json','NegativeResolutionPairGate.json',
 'PriorDirectedRegressionProof.json','SourceLargerHPairProof.json','Run-TestHMixedStock.ps1','Run-TestRealColdReplay.ps1',
 'Run-TestOperationalArcRegressions.ps1','Run-TestSpatialSteelRawControl.ps1','Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1',
 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1','tools/build_workbook/Run-Audit02ModeTests.ps1',
 'PreparedSpatialRawBeforeCommentAlignment.bas','PreparedCold56PairsBeforeCommentAlignment.bas','PreparedSequentialHBeforeCommentAlignment.bas']
raw={}
for number,root in roots.items():
    names=fixed+[p.relative_to(root).as_posix() for directory in ('src','tests','Native') for p in (root/directory).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    names+=['Input/'+n for n in ('Reverse1.bin','SourceColdRawHistory.bin','EnvelopeOptions.bin','NativeTolerance.bin','NativeSolver.json')]
    for name in sorted(set(names)):
        if (root/name).is_file():raw[f'{number}/{name}']=(root/name).read_bytes()
    if number in (885,886,888,891):
        with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:raw[f'{number}/ActualCompiledVBA.bin']=book.read('xl/vbaProject.bin')
with ZipFile(roots[890]/'SavedWorkbook.xlsm') as book:raw['890/ActualSavedVBA.bin']=book.read('xl/vbaProject.bin')
raw['Referenced835/CCoupledElasticState.cls']=(base/'Cold835GatedColdReplay_20261125/src/Solver/CCoupledElasticState.cls').read_bytes()
raw['Referenced835/EnvelopeEnergy.Accepted.Operators.csv']=(base/'Cold835GatedColdReplay_20261125/Native/EnvelopeEnergy.Accepted.Operators.csv').read_bytes()
raw['Referenced880/Stock.RawHistory.bin']=(base/'Cold880CriticalHCold96_20261126/Native/Stock.RawHistory.bin').read_bytes()
pins={n:sha256(b).hexdigest().upper() for n,b in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=9) as bundle:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);bundle.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as bundle:
    for name,digest in pins.items():assert sha256(bundle.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=['Stage-TestSpatialSteelRawControl.py','Check-TestSpatialSteelRawControl.py','Align-TestSpatialRawComments.py',
 'Stage-TestCold56LargerHPairs.py','Check-TestCold56LargerHPairs.py','Align-TestCold56PairComments.py',
 'Stage-TestVerifiedEnvelopeGrowth.py','Stage-TestStrictEnvelopeGrowth.py','Stage-TestPreparedRegressionRun.py',
 'Stage-TestStrictGrowthCold96.py','Check-TestStrictGrowthCold96.py','Check-TestCriticalHCold96.py','Check-TestFixedMetricCold.py',
 'Check-TestSpatialProjectionCompatibility.py','Stage-TestStrictGrowthColdStress.py','Check-TestStrictGrowthColdStress.py',
 'Stage-TestImmutableColdStressReplay.py','Check-TestCriticalHColdStress.py','Check-TestVerifiedColdStressReplay.py',
 'Stage-TestSequentialAcceptedH.py','Check-TestSequentialAcceptedH.py','Align-TestSequentialHComments.py',
 'Check-TestSharedCurveHControl.py','Check-TestVerifiedHCornerPairs.py','Check-TestPersistentHEvidence.py',
 'Read-TestColdRawCapture.py','Stage-TestColdRawReplay.py','Check-TestActualReverseCapture.py','Check-TestHardHalfSteps.py',
 'Validate-TestHardPrepared.py','Stage-TestEnergyNumerics.py',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
summary=dict(independent_steel_boundary=steel,cold56_H_pairs=pairs,sequential_accepted_H=seq,regressions889=reg,
 cold890=cold,cold892_stress=stress,actual_cold_comparison=comparison,raw_archive=archive.name,raw_files=len(raw),
 raw_bytes=sum(map(len,raw.values())),xz_preset=9,free_physical_before=memory.free_physical,free_commit_before=memory.free_page,
 full_acceptance_not_run=True,physical_capacity_certified=False,no_release_or_main_promotion=True,whole_goal_complete=False)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes());assert not state['active_jobs']
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
 latest_verified_growth_history_evidence=out.relative_to(repo).as_posix(),latest_independent_steel_boundary=steel,
 latest_cold56_larger_H_pairs=pairs,latest_sequential_accepted_H=seq,latest_strict_growth_regressions=reg,
 latest_completed_cold890=cold,latest_completed_cold890_stress=stress,latest_actual_cold890_comparison=comparison,
 active_cold_source_root=None,active_cold_exec_session=None,
 previous_goal_turn_classification='progress:885native14PASS six raw equilibria independently rebuilt; cold-zero separate Newton finds steel side boundary1.6392967139708858 in3iterations within native brackets.886actualcold56 strictH ratios1,2,4,8 confirmed2/2/2/3iterations,16/32unconfirmed; full32root alone not admitted due half16.888conditional pairedHgrowth2 compiled;887nearresolution bypass counterexample preserved/fixed.889all164PASS.890actualcold complete731.138671875s/directOK/capacityNumFail/max1.1042833676987807,all56fields/fractions/lambdas bitidentical880,acceleratedHattempts0,actualfinalcontrol signedspatialarc.892all56stresses/sixforces/56+4/all3steel verified.891full8 versus accepted4+4 difference1.619461e-4increment/lambda4.6e-9; next accepted8+8unconfirmed. No physicalultimate/release/main promotion.',
 next_action='Target the actual signed-arc continuation at cold state56 and earlier state19: explain why valid strict H pairs are not selected on actual path, compare Arc radius with H-guide desired scale and TryEvent eligibility. Inspect exact source before asserting H-handoff bug: H-guide currently seeds a material-event search and is not automatically an H-controlled candidate. Do not repeat cold run for unused H-growth change. Preserve cold890 and sequential891 evidence; independent full internal material-state vectors and full SOW/physical ultimate remain open.')
data=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for file in (path,here/'Continuation666_Resume20261123.json'):file.write_bytes(data)
print('ARCHIVED',len(raw),'RAW FILES',archive.stat().st_size,'BYTES;COLD890 EXACT880;NO H-GROWTH USED;STEEL/SEQUENTIAL PROOFS VERIFIED;GOAL ACTIVE')
