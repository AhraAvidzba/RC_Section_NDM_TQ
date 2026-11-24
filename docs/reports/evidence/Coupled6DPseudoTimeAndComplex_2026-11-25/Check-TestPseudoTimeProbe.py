"""ДЛЯ ТЕСТОВ. Проверяет фактические направления739, их радиус и отрицательный исход без подмены тайм-аута результатом НДС."""
from pathlib import Path
from runpy import run_path
import json,csv,re
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold739PseudoTimeProbe_20261125';parent=base/'Cold738TailFullDifferences_20261125'
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackReplay')
assert native['timeout600'] and native['own_process_closed']
outcome=(root/'Native/PostCrackRawReplay.PseudoOutcome.txt').read_text('cp1251')
assert 'PTC_RESULT:' not in outcome and 'PTC_CANDIDATE: improved=True;' in outcome
for suffix in ('.csv','.Operators.csv'):
    name='Native/PostCrackRawReplay.Accepted'+suffix
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteResponse.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
data=(root/'Native/PostCrackRawReplay.TailMatrix2.bin').read_bytes();assert data==(parent/'Native/PostCrackRawReplay.TailMatrix2.bin').read_bytes()
reader=Reader(data);size,ready=reader.ReadTest('<2i');ratio,factor=reader.ReadTest('<2d')
seed=reader.VectorTest();scales=reader.VectorTest();baseline=reader.VectorTest();a=np.array(reader.ReadTest('<3721d')).reshape(size,size)
assert ready and size==61 and ratio==1.25
data=(root/'Native/PostCrackRawReplay.PseudoReference.bin').read_bytes();assert data[:16]==b'NDMPTCREF0000001'
reader=Reader(data);n,=reader.ReadTest('<i');accepted_coordinates=reader.VectorTest();reader.VectorTest()
reference=np.array(reader.ReadTest('<3721d')).reshape(size,size);assert n==size and reader.position==len(data)
old=Reader((base/'Cold737TailMatrixPathRetry_20261125/Native/PostCrackRawReplay.TailMatrix0.bin').read_bytes())
old.ReadTest('<2i');old.ReadTest('<2d');accepted=old.VectorTest();old.VectorTest();old.VectorTest()
ref0=np.array(old.ReadTest('<3721d')).reshape(size,size)
assert np.array_equal(reference,ref0) and np.array_equal(seed+scales[:60]*accepted_coordinates[:60],accepted)

rows=list(csv.reader((root/'Native/PostCrackRawReplay.PseudoTrials.csv').read_text('cp1251').splitlines(),delimiter=';'))
base_row,=[x for x in rows if x[0]=='BASE'];reported_initial=float(base_row[1]);reported_radius=float(base_row[2])
linear={float(x[1]):x for x in rows if x[0]=='LINEAR'}
trials=[x for x in rows if x[0]=='TRIAL'];assert len(trials)==12 and len(linear)==4
stresses={(float(x[1]),int(x[2])):np.array(list(map(float,x[3:]))) for x in rows if x[0]=='RESIDUAL'}
assert len(stresses)==12 and all(x[2]=='True' for x in linear.values())

points=26;dofs=60;physical=56
b=np.zeros((points,6,dofs));g=b.copy();areas=np.zeros(points)
for kind,*v in run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows'](root/'Native/PostCrackRawReplay.Accepted.Operators.csv'):
    if kind=='B':
        p,d=int(v[0])-1,int(v[1])-1;b[p,:,d]=np.array(v[2:8],float);g[p,:,d]=np.array(v[8:],float)
    elif kind=='P':areas[int(v[0])-1]=float(v[3])
weights=np.array([1.,1.,1.,.5,.5,.5])

def TestDistance(delta):
    """Повторяет подтверждённую пространственную метрику трёх станций без служебных множителей."""
    movement=np.zeros(dofs);movement[:physical]=delta[:physical]
    center=np.einsum('pcd,d->pc',b,movement);slope=np.einsum('pcd,d->pc',g,movement)
    strains=np.array([center+s*.001*slope for s in (-1,0,1)])
    return float(np.sqrt(np.einsum('spc,p,c->',strains**2,areas,weights)/(3*sum(areas))))

radius=TestDistance(seed-accepted)
assert abs(radius-reported_radius)<=1e-11*radius and abs(reported_initial-max(abs(baseline)))<=64*np.finfo(float).eps*reported_initial
load_column=a[:,60]

def TestBalancedNorm(values):
    """Диагностически исключает линейный множитель нагрузки; это анализ сохранённых данных, не дополнительное решение НДС."""
    corrected=values-load_column*((load_column@values)/(load_column@load_column))
    return float(max(abs(corrected)))

results=[]
for row in trials:
    mu,cut,gain=float(row[1]),int(row[2]),float(row[3]);assert row[4]=='True'
    coordinates=np.array(list(map(float,row[6:])));assert len(coordinates)==61 and coordinates[0]==0
    values=stresses[mu,cut];assert len(values)==61
    assert abs(max(abs(values))-float(row[5]))<=64*np.finfo(float).eps*max(abs(values))
    raw_step=coordinates/gain;matrix=a+mu*reference
    defect=float(max(abs(matrix@raw_step+baseline)))
    backward=defect/(np.linalg.norm(matrix,ord=np.inf)*max(abs(raw_step))+max(abs(baseline)))
    assert backward<=100*np.finfo(float).eps*size
    raw_field=seed+scales[:60]*raw_step[:60];length=TestDistance(raw_field-seed)
    expected_gain=min(1.,radius/length)/(2**cut) if length else 1/(2**cut)
    assert abs(gain-expected_gain)<=1e-11*expected_gain
    trial_field=seed+scales[:60]*coordinates[:60];distance=TestDistance(trial_field-seed)
    assert distance<=radius*(1+1e-11)
    results.append(dict(mu=mu,cut=cut,gain=gain,native_norm=float(max(abs(values))),
        load_balanced_norm=TestBalancedNorm(values),spatial_distance=distance,spatial_radius=radius,
        independent_linear_backward_error=backward,native_linear_minimum_pivot=float(linear[mu][3])))
best=min(p['native_norm'] for p in results);initial=max(abs(baseline))
proof=dict(native=native,numerical_core_exact738=True,all70_tensor_fields_exact_verified_parent=True,
    fresh_native_reference_matrix_exact737_accepted_matrix=True,accepted_field_roundtrip_exact=True,
    reference_rank=61,source_rank=17,readonly_trials=results,all12_available=True,
    all12_spatial_caps_and_linear_backward_errors_verified=True,
    initial_norm=float(initial),best_norm=best,relative_norm_reduction=float(1-best/initial),
    initial_load_balanced_norm=TestBalancedNorm(baseline),
    best_trial_load_balanced_norm=min(p['load_balanced_norm'] for p in results),
    ordinary_root_verification_completed=False,numerical_stop_not_a_returned_NumFail=True,
    final_NDS_result_missing=True,actual_trial_vectors_recorded_as_native15digit_CSV_not_full_binary=True,
    no_working_promotion=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'PseudoTimeProbeProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('739 ALL12 DIRECTION/CAP CHECKS VERIFIED; NORM REDUCTION',proof['relative_norm_reduction'],
    'ROOT VERIFICATION TIMEOUT600; NO NDS RESULT; BALANCED RESIDUAL NOT IMPROVED',proof['best_trial_load_balanced_norm']>=proof['initial_load_balanced_norm'])
