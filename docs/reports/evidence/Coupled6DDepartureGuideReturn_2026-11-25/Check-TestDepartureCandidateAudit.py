"""ДЛЯ ТЕСТОВ. Независимо проверяет НДС материального кандидата749 и направление его механического приращения."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold749DepartureCandidateAudit_20261125';source=base/'Cold747EarlyEnvelopeFrozenReplay_20261125'
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native['passed']==1 and native['failed']==0 and not native['timeout600']
assert (root/'Input/Reverse1.bin').read_bytes()==(source/'Input/Reverse1.bin').read_bytes()
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
assert len(original['fields'])==46
files=list((root/'Native').glob('Departure.Candidate*.bin'));assert len(files)==1
raw=files[0].read_bytes();assert raw[:16]==b'NDMDEPARTCAND001'
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];reader=Reader(raw)
context=reader.VectorTest();accepted=reader.VectorTest();candidate=reader.VectorTest()
projection=reader.VectorTest();old_projection=reader.VectorTest()
assert reader.position==len(raw) and len(context)==9
factor,reference,uncertainty,precision_ready,tolerance,norm,old_target,revision,modulus=context
assert precision_ready==1 and norm<=tolerance and reference==original['lambdas'][-1]
assert np.array_equal(accepted,original['fields'][-1])
mod=run_path(str(here/'Validate-TestHardPrepared.py'))
mechanics,fields=mod['TestValidate'](root,'Departure.Sequence')
assert len(fields)==47 and mechanics['details'][0]=='DiagnosticReSolve'
np.testing.assert_allclose(fields[:-1],original['fields'],rtol=1e-14,atol=1e-300)
np.testing.assert_allclose(fields[-1],candidate,rtol=1e-14,atol=1e-300)
rows=mod['TestRows'](root/'Native/Departure.Sequence.Operators.csv')
dimension=next(r for r in mod['TestRows'](root/'Native/Departure.Sequence.csv') if r[0]=='DIM')
dofs,physical,points,concrete=map(int,dimension[1:5]);half=float(dimension[5])
b=np.zeros((points,6,dofs));gradient=b.copy();areas=np.zeros(points)
for kind,*r in rows:
    if kind=='B':p,d=int(r[0])-1,int(r[1])-1;b[p,:,d]=np.asarray(r[2:8],float);gradient[p,:,d]=np.asarray(r[8:],float)
    elif kind=='P':areas[int(r[0])-1]=float(r[3])
weights=np.array([1.,1.,1.,.5,.5,.5])
def TestTensor(delta):
    """Исключает множители связей и восстанавливает три станции принятой механической метрики."""
    x=np.zeros(dofs);x[:physical]=delta[:physical]
    return np.asarray([np.einsum('pcd,d->pc',b+s*half*gradient,x) for s in (-1,0,1)])
def TestDot(x,y):
    """Считает среднее скалярное произведение инженерных деформаций с половинным весом сдвигов."""
    return float(np.einsum('spc,spc,p,c->',x,y,areas,weights)/(3*sum(areas)))
before=TestTensor(accepted-original['fields'][-2]);move=TestTensor(candidate-accepted)
length=float(np.sqrt(TestDot(before,before)));distance=float(np.sqrt(TestDot(move,move)))
advance=TestDot(before,move)/length
old_delta=float(old_projection@(candidate-accepted));new_delta=float(projection@(candidate-accepted))
assert advance<0 and old_delta<0 and new_delta>0
assert abs(advance-old_delta)<=1e-11*max(abs(advance),abs(old_delta))
native_outcome=(root/'Native/Departure.Outcome.txt').read_text('cp1251')
assert native_outcome.startswith('SOLVED;0;PROBES;1;LAMBDA;0;CANDIDATES;1')
assert f'ACCEPTED_REVISION_UNCHANGED;{int(revision)}' in native_outcome
load_resolution=8*float(original['options'][4])*(1+abs(reference))
proof=dict(native=native,source_states=46,diagnostic_sequence_states=47,diagnostic_candidate_not_accepted_by_controller=True,
    candidate_raw_sha256=sha256(raw).hexdigest().upper(),candidate_lambda=float(factor),reference_lambda=float(reference),
    lambda_change=float(factor-reference),lambda_uncertainty=float(uncertainty),lambda_uncertainty_available=True,
    native_full_residual_MPa=float(norm),native_full_tolerance_MPa=float(tolerance),
    independent_candidate_equilibrium=mechanics['states'][-1],all46_original_states_equilibrium_verified_again=True,
    independent_spatial_signed_advance=float(advance),previous_spatial_increment=length,spatial_candidate_distance=distance,
    old_projection_increment=old_delta,new_material_projection_increment=new_delta,
    original_requested_control=float(old_target),candidate_original_control=float(old_projection@candidate),
    lambda_resolution=float(load_resolution),unchanged_existing_lambda_guard_would_reject=True,
    independent_forward_direction_check_would_also_reject=True,accepted_revision_unchanged=int(revision),
    repeated_working_procedure_from_fresh_exact_history_not_claimed_original_call_bitidentity=True,
    removing_lambda_guard_alone_not_supported_by_this_evidence=True,no_new_cold_capacity_search=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'DepartureCandidateAuditProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('749 CANDIDATE NDS INDEPENDENTLY VERIFIED; LAMBDA',factor,'FROM',reference,
    'SIGNED FIELD ADVANCE',advance,'BACKWARD: DO NOT REMOVE ACCEPTANCE GUARDS')
