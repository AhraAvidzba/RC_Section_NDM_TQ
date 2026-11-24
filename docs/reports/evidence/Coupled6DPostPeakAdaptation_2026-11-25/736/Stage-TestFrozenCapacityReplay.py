"""ДЛЯ ТЕСТОВ. Готовит однопроходную проверку точной истории завершённого холодного поиска."""
from argparse import ArgumentParser
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re,struct
import numpy as np

p=ArgumentParser(description=__doc__);p.add_argument('--root',required=True,type=Path);p.add_argument('--parent',required=True,type=Path)
a=p.parse_args();root=a.root;parent=a.parent;here=Path(__file__).resolve().parent
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](parent,'PostCrackTension')
assert native['own_process_closed']
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](parent/'Native/Stock.RawHistory.bin')
assert raw['incomplete_tail_bytes']==0
frame=frames[-1];count=len(frame['fields']);assert count>=2
template=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Native/PostCrackPrescribed.Accepted2.bin')
for key in ('fields','fractions','lambdas'):assert np.array_equal(frame[key][:2],template[key]),key
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'FrozenReplayCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())

def TestVector(v):
    """Сохраняет исходные двоичные Double, не текстовое округление."""
    b=np.asarray(v,dtype='<f8');return struct.pack('<i',len(b))+b.tobytes()

data=bytearray(b'NDMPOSTREPLAY001'+struct.pack('<i',60))
for key in ('calibration','options','offset','direction'):data+=TestVector(template[key])
data+=struct.pack('<i',count)
for fraction,factor,field in zip(frame['fractions'],frame['lambdas'],frame['fields']):
    data+=struct.pack('<dd',fraction,factor)+TestVector(field)
values['Input/PostCrackReplay.bin']=bytes(data)
values['Input/PostCrackReplayOptions.bin']=struct.pack('<i3d',3,0,0,0)
values['Input/PostCrackSourceRawHistory.bin']=(parent/'Native/Stock.RawHistory.bin').read_bytes()
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Friend Function ReplayAcceptedFields\(.*?^End Function',code)
anchor='        mAcceptedRevision = mAcceptedRevision + 1\n';assert method.count(anchor)==1
hook='        modTestCoupledElasticState.TestObserveFrozenReplay index, field\n'
assert hook not in code
changed=method.replace(anchor,anchor+hook,1);assert changed.replace(hook,'')==method
values[name]=code.replace(method,changed,1).replace('\n','\r\n').encode('utf-8-sig')
reference=parent.parent/'Cold733FastFrozenReplay_20261124'
values['tests/modTestCoupledElasticState.bas']=(reference/'tests/modTestCoupledElasticState.bas').read_bytes()
values['FrozenObserverSourceAudit.json']=json.dumps(dict(source=str(reference),source_proof_sha256=sha256((reference/'FastFrozenReplayProof.json').read_bytes()).hexdigest().upper(),
    original_compilation_failure_repaired_before_reuse=True,only_mode0_frozen_replay=True),indent=2).encode('utf-8')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledNewtonSolver','modTestCoupledElasticState')")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['Run-TestPostCrackRawReplay.ps1']=(reference/'Run-TestPostCrackRawReplay.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,b in values.items():(root/name).write_bytes(b);prep['inputs'][name]=sha256(b).hexdigest().upper()
prep['directed_scope'].update(case='FrozenCapacitySinglePassStressReplay',actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],
    source_root=str(parent),source_states=count,source_native_search_timeout=native['timeout600'],original_binary_fields_and_loads_preserved=True,
    numerical_methods_exact_except_readonly_replay_observer=True,original_force_gauge_and_steel_checks=True,
    full_independent_field_validation_required=True,no_new_capacity_search=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('FROZEN REPLAY STAGED:',count,'ORIGINAL STATES; NO NEW SEARCH')
