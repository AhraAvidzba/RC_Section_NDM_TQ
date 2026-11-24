"""ДЛЯ ТЕСТОВ. Восстанавливает12 точных полей настоящего холодного741 на его материале для независимого интегрирования."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold742RealColdFrozenReplay_20261125';parent=base/'Cold741OrdinaryFirstComplex_20261125'
proof=json.loads((parent/'OrdinaryFirstComplexProof.json').read_bytes())
assert proof['cold741']['capacity']=='NumFail' and proof['cold741']['native']['own_process_closed']
read=run_path(str(here/'Read-TestColdRawCapture.py'));frames,raw=read['TestHistory'](parent/'Native/Stock.RawHistory.bin')
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
frame=frames[-1];assert len(frame['fields'])==12 and raw['incomplete_tail_bytes']==0
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'FrozenCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
values['Input/Reverse1.bin']=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,frame)
values['Input/Source741History.bin']=(parent/'Native/Stock.RawHistory.bin').read_bytes()
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
pattern=r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub';method,=re.findall(pattern,code)
end=method.index('    TestExportReverseCandidate 0\n')
changed=method[:end]+'''    mPtcSuppressTest = True
    TestExportFastFrozenReplay
    mPtcSuppressTest = False
    mPolishExecutedTest = True
End Sub'''
code=code.replace(method,changed,1)
assert max(map(len,code.splitlines()))<1024
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=(here/'Compile-TestExternalPoint.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
runner=(here/'Run-TestPostCrackRawReplay.ps1').read_text('utf-8-sig') if (here/'Run-TestPostCrackRawReplay.ps1').is_file() else (parent/'Run-TestPostCrackRawReplay.ps1').read_text('utf-8-sig')
runner=runner.replace('modTestCoupledElasticState.RunTestPostCrackRawReplay','modTestCoupledElasticState.RunTestReversePolish')
runner=runner.replace('PostCrackReplay.txt','RealColdReplay.txt').replace('-TimeoutSeconds 600','-TimeoutSeconds 180')
values['Run-TestRealColdReplay.ps1']=runner.replace('\n','\r\n').encode('utf-8-sig')
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ActualCold741Frozen12',actual_import_this_pass=['modTestCoupledElasticState'],
    all_numerical_core_exact741=True,source_raw_history_exact=True,source_states=12,
    one_stress_only_replay=True,actual_six_nonzero_loads=True,pseudo_time_not_called=True,
    no_new_NDS_or_capacity_search=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('742 EXACT12 REAL COLD FIELDS STAGED; NUMERICAL CORE EXACT741; NO NEW CAPACITY SEARCH')
