"""ДЛЯ ТЕСТОВ. Собирает нативное восстановление точной холодной истории нового рабочего решателя; поиск НДС и предела не запускается."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from argparse import ArgumentParser
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__)
p.add_argument('--source',default='Cold794PreEventBudget_20261125');p.add_argument('--root',default='Cold798AdaptiveColdReplay_20261125')
args=p.parse_args();source=base/args.source;root=base/args.root
outcome=json.loads((source/'OperationalColdProof.json').read_bytes());assert outcome['cold']['native']['own_process_closed']
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](source/'Native/Stock.RawHistory.bin');assert raw['incomplete_tail_bytes']==0
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](source/'Input/Reverse1.bin')
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,source,1e-6,1,1e-6,0)
archive=root/'AdaptiveColdReplayCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchors=('    TestRunOperationalArc\n','    TestRunAutomaticLoadingWalk\n')
assert sum(code.count(anchor) for anchor in anchors)==1
for anchor in anchors:code=code.replace(anchor,'',1)
for anchor in ('    TestRecentLoadingDirectionGuards\n','    TestRecentLoadingResetGuard\n'):
    assert code.count(anchor)<=1;code=code.replace(anchor,'',1)
values={name:code.replace('\n','\r\n').encode('utf-8-sig'),
    'Input/Reverse1.bin':run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,frames[-1]),
    'Input/SourceColdRawHistory.bin':(source/'Native/Stock.RawHistory.bin').read_bytes(),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(source/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(source/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ActualAdaptiveColdStressReplay',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    source_root=source.name,source_states=len(frames[-1]['fields']),all_source_cold_fields_exact=True,
    all_working_numerical_methods_exact_source=True,no_new_NDS_search=True,controller_not_called=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root.name,'EXACT COLD STATES',len(frames[-1]['fields']),'FROM',source.name,'STRESS ONLY; ALL WORKING CODE EXACT')
