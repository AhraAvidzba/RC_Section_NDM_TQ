"""ДЛЯ ТЕСТОВ. Готовит только повтор материальных откликов точной завершённой холодной истории без нового поиска НДС."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from argparse import ArgumentParser
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--source',required=True);p.add_argument('--root',required=True);p.add_argument('--proof',required=True)
args=p.parse_args();source=base/args.source;root=base/args.root
outcome=json.loads((source/args.proof).read_bytes());assert outcome['cold']['native']['own_process_closed']
assert not outcome['cold'].get('workflow_result_missing',False)
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](source/'Native/Stock.RawHistory.bin')
assert raw['incomplete_tail_bytes']==0 and outcome['accepted_states']==len(frames[-1]['fields'])
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](source/'Input/Reverse1.bin')
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,source,1e-6,1,1e-6,0)
archive=root/'VerifiedColdStressCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub',code)
start=method.index('    mPtcSuppressTest = False\n')+len('    mPtcSuppressTest = False\n')
end=method.index('    mPolishExecutedTest = True\n',start)
removed=method[start:end];assert re.fullmatch(r'(?:    Test\w+\n)+',removed),removed
updated=method[:start]+method[end:];code=code.replace(method,updated,1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values['Input/Reverse1.bin']=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,frames[-1])
values['Input/SourceColdRawHistory.bin']=(source/'Native/Stock.RawHistory.bin').read_bytes()
values.update({'SourceColdProof.json':(source/args.proof).read_bytes(),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(source/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(source/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='VerifiedColdStressOnlyReplay',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    source_root=source.name,source_proof=args.proof,source_states=len(frames[-1]['fields']),source_history_exact=True,
    all_working_numerical_and_material_methods_exact_source=True,only_diagnostic_calls_after_frozen_export_removed=removed.splitlines(),
    no_new_NDS_or_capacity_search=True,controller_not_called=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'EXACT COLD STATES',len(frames[-1]['fields']),'SOURCE',source.name,'STRESS REPLAY ONLY')
