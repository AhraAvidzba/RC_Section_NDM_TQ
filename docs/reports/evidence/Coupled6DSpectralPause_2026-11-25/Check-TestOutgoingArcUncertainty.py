"""ДЛЯ ТЕСТОВ. Независимо проверяет исходное равновесие и сопряжённую оценку чувствительности λ, не выдавая линейную оценку за физический предел."""
from pathlib import Path
from runpy import run_path
import json,struct
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold829OutgoingArcUncertainty_20261125';parent=base/'Cold826AutomaticColdReplay_20261125';source=base/'Cold825ExtendedLoadingWalk_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledDeformationContinuation.cls',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
source_input=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](source/'Input/Reverse1.bin')
actual=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
fields=list(source_input['fields']);lambdas=list(source_input['lambdas'])
history=json.loads((root/'Input/Source825Attempts.json').read_bytes())
for attempt in history['attempts']:
    if not attempt['solved']:continue
    r=Reader((source/'Native'/('LoadingWalk.'+str(attempt['attempt']).zfill(3)+'.bin')).read_bytes())
    assert r.raw[:16]==b'NDMOPERARC000004'
    packet=[r.VectorTest() for _ in range(15)];field,forces,projection=[r.VectorTest() for _ in range(3)]
    r.VectorTest();r.VectorTest();r.VectorTest();assert r.position==len(r.raw)
    fields.append(field);lambdas.append(packet[0][23])
assert len(fields)==69 and np.array_equal(fields,actual['fields']) and np.array_equal(lambdas,actual['lambdas'])
r=Reader((root/'Native/OutgoingUncertainty.Result.bin').read_bytes());assert r.raw[:16]==b'NDMARCUNCERT0001\0'[:16]
meta,field,projection,before,after=[r.VectorTest() for _ in range(5)];assert r.position==len(r.raw)
assert len(meta)==11 and meta[0]==69 and meta[3]==68 and meta[4]==meta[5] and meta[6]==1 and meta[9]==0
assert meta[1]==lambdas[-1] and meta[2]==lambdas[-2] and np.array_equal(field,fields[-1])
assert before[0]==0 and before[5]==1 and before[6]==after[6]==meta[4]
validate=run_path(str(here/'Validate-TestHardPrepared.py'))
mechanics,csv_fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(csv_fields,fields,rtol=1e-14,atol=1e-300)
assessment=None;failure_stage=None
if meta[10]:
    assert meta[10]==1
    A,rhs,residual,factor,tolerance=run_path(str(here/'Check-TestActualAdjointResidual.py'))['TestRead'](root/'Native/Adjoint.1.bin')
    assert factor==meta[1] and np.array_equal(rhs,np.r_[np.zeros(60),1.])
    outcomes=validate['TestRows'](root/'Native/Adjoint.Outcomes.csv');assert len(outcomes)==1 and outcomes[0][0]=='1'
    solved=outcomes[0][1]=='1';assert solved==bool(meta[7])
    col=np.max(abs(A),axis=0);col[col==0]=1
    scaled=A/col;row=np.max(abs(scaled),axis=1);row[row==0]=1
    scaled=scaled/row[:,None];u,s,vh=np.linalg.svd(scaled,full_matrices=True)
    ranks=[]
    for threshold in (np.finfo(float).eps*61*s[0],1e-12*s[0]):
        rank=int(np.sum(s>threshold));target=rhs/row
        y=vh[:rank].T@((u[:,:rank].T@target)/s[:rank]);weights=y/col
        ranks.append(dict(threshold=float(threshold),rank=rank,scaled_equation_defect=float(np.max(abs(scaled@y-target))),
            physical_adjoint_defect=float(np.max(abs(A@weights-rhs)))))
    assessment=dict(available=solved,lambda_value=float(factor),force_bound_MPa=float(tolerance),
        actual_scaled_residual=float(np.max(abs(residual))),scaled_singular_values=s.tolist(),rank_diagnostics=ranks)
    if solved:
        raw=(root/'Native/Adjoint.Weights.1.bin').read_bytes();size,=struct.unpack_from('<i',raw)
        assert size==61 and len(raw)==4+61*8
        weights=np.frombuffer(raw,dtype='<f8',offset=4).copy()
        defect=A@weights-rhs
        backward=float(np.max(abs(defect))/(np.linalg.norm(A,np.inf)*np.max(abs(weights))+np.max(abs(rhs))))
        assert backward<=128*61*np.finfo(float).eps
        floor=2e-12*(1+abs(factor));norm=max(tolerance,float(np.max(abs(residual))))
        uniform=float(np.sum(abs(weights))*norm+floor)
        componentwise=float(np.sum(abs(weights*residual))+floor)
        signed=float(abs(weights@residual)+floor)
        np.testing.assert_allclose(meta[8],uniform,rtol=1e-12,atol=1e-15)
        assessment.update(adjoint_backward_error=backward,weights_l1=float(np.sum(abs(weights))),
            uniform_linearized_uncertainty=uniform,actual_componentwise_linearized_estimate=componentwise,
            actual_signed_linearized_estimate=signed,
            observed_source825_load_decline=float(max(lambdas)-lambdas[-1]),
            nonlinear_remainder_not_bounded=True)
    else:failure_stage='AdjointLU'
else:
    assert not meta[7] and not (root/'Native/Adjoint.1.bin').exists()
    failure_stage='MaterialTangentBeforeAdjoint'
proof=dict(native=native,all_actual_source69_fields_exact825=True,all_working_methods_exact826=True,
    source69_equilibria_independently_rebuilt=mechanics,meta=meta.tolist(),before_effort=before.tolist(),after_effort=after.tolist(),
    exact_known_root_restored_without_correction=True,no_new_material_history_accepted=True,
    assessment=assessment,unavailable_stage=failure_stage,
    local_linear_sensitivity_only=True,nonlinear_remainder_not_certified=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'OutgoingArcUncertaintyProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('829 LAMBDA SENSITIVITY',assessment,'UNAVAILABLE STAGE',failure_stage,'NO PHYSICAL CAPACITY CERTIFIED')
