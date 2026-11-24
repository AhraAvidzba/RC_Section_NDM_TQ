"""ДЛЯ ТЕСТОВ. Сопоставляет полную H-производную настоящего шестикомпонентного сечения и проверяет неизменность истории."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,struct,numpy as np
from argparse import ArgumentParser

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,default=base/'Cold746ActualHFullDifferences_20261125')
root=parser.parse_args().root;parent=base/'Cold744RecoveredColdFrozenReplay_20261125'
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native['passed']==1 and native['failed']==0 and not native['timeout600']
for suffix in ('.csv','.Operators.csv'):
    name='Native/EnvelopeEnergy.Accepted'+suffix
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledNewtonSolver.cls',
             'src/Solver/CCoupledDeformationContinuation.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
data=(root/'Native/ActualH32.Matrix.bin').read_bytes();assert data[:16]==b'NDMTAILMATRIX001'
reader=Reader(data);size,available=reader.ReadTest('<2i');ratio,factor=reader.ReadTest('<2d')
seed=reader.VectorTest();scales=reader.VectorTest();assert size==61 and available and ratio==1.25
r=reader.VectorTest();a=np.asarray(reader.ReadTest('<'+str(size*size)+'d')).reshape(size,size)
assert reader.position==len(data)
context=(root/'Native/ActualH32.Context.bin').read_bytes();length,=struct.unpack_from('<i',context)
assert length==9 and len(context)==4+8*length
c=np.frombuffer(context,dtype='<f8',offset=4).copy();before,after,target,baseline,resolution,row_scale,last,prior,revision=c
assert after==baseline and after>before and target==after+ratio*(after-before)
assert row_scale==10000. and factor==last+ratio*(last-prior)
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
assert len(original['fields'])==32 and last==original['lambdas'][-1] and prior==original['lambdas'][-2]
expected=original['fields'][-1].copy();expected[:56]+=ratio*(expected[:56]-original['fields'][-2,:56])
assert np.array_equal(seed,expected)
fd=run_path(str(here/'Check-TestTailFullDifferences.py'))
n,expected_count,assembled,samples,raw=fd['TestReadSamples'](root/'Native/ActualH32.TailDifferences.bin')
assert n==size and expected_count==247 and np.array_equal(assembled,r)
base_sample=samples[0];assert base_sample['ready'] and base_sample['step']==0 and base_sample['column']==0
base_residual=base_sample['residual'];steps=[];matrices={}
for h in (1e-6,1e-4):
    group=[s for s in samples if s['step']==h];assert len(group)==2*size+1
    repeated,=[s for s in group if s['column']==0]
    missing=[dict(column=s['column'],sign=s['sign']) for s in group if not s['ready']]
    item=dict(step=h,unavailable=missing,base_repeated_exactly=bool(repeated['ready'] and np.array_equal(repeated['residual'],base_residual)))
    if not missing:
        plus=np.column_stack([next(s['residual'] for s in group if s['column']==j and s['sign']==1) for j in range(1,size+1)])
        minus=np.column_stack([next(s['residual'] for s in group if s['column']==j and s['sign']==-1) for j in range(1,size+1)])
        central=(plus-minus)/(2*h);forward=(plus-base_residual[:,None])/h;backward=(base_residual[:,None]-minus)/h
        matrices[h]=central
        item.update(central=fd['TestSpectrum'](central),forward=fd['TestSpectrum'](forward),backward=fd['TestSpectrum'](backward),
            relative_frobenius_difference_to_supplied=float(np.linalg.norm(central-a)/(np.linalg.norm(central)+np.finfo(float).tiny)),
            maximum_abs_difference_to_supplied=float(np.max(abs(central-a))),
            max_difference_row_1based=int(np.unravel_index(np.argmax(abs(central-a)),a.shape)[0]+1),
            max_difference_column_1based=int(np.unravel_index(np.argmax(abs(central-a)),a.shape)[1]+1),
            forward_backward_abs_difference_max=float(np.max(abs(forward-backward))))
    steps.append(item)
comparison=None
if len(matrices)==2:
    comparison=dict(relative_frobenius_difference=float(np.linalg.norm(matrices[1e-6]-matrices[1e-4])/np.linalg.norm(matrices[1e-4])),
        maximum_abs_difference=float(np.max(abs(matrices[1e-6]-matrices[1e-4]))))
proof=dict(native=native,real_six_component_source_history_states=32,accepted_tensor_output_exact_verified744=True,
    numerical_kernel_exact744=True,readonly_secant_ratio=ratio,new_forecast_not_exact_failed_request=True,
    context=dict(before_H=before,after_H=after,target_H=target,resolution=resolution,row_scale=row_scale,
        accepted_lambda=last,previous_lambda=prior,seed_lambda=factor,accepted_revision=revision),
    original_binary_seed_matches_physical_secant_exactly=True,assembled=fd['TestSpectrum'](a),
    assembled_vs_pure_residual_bitidentical=bool(np.array_equal(r,base_residual)),
    assembled_vs_pure_residual_max_difference=float(np.max(abs(r-base_residual))),
    source_matrix_sha256=sha256(data).hexdigest().upper(),samples=raw,sample_count=len(samples),steps=steps,
    finite_difference_step_comparison=comparison,no_new_NDS_or_capacity_search=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'ActualHFullDifferenceProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('746 ACTUAL ALL6 H MATRIX RANK',proof['assembled']['rank'],'VALUE/JAC IDENTICAL',proof['assembled_vs_pure_residual_bitidentical'])
for step in steps:print('FD',step['step'],'MISSING',len(step['unavailable']),'RANK',step.get('central',{}).get('rank'),
    'RELATIVE_DIFFERENCE',step.get('relative_frobenius_difference_to_supplied'))
