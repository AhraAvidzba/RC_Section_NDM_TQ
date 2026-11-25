"""ДЛЯ ТЕСТОВ. Отличает возврат к исходной цели от нового вспомогательного корня и проверяет общий30-итерационный бюджет."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold750DepartureGuideReturn_20261125';parent=base/'Cold749DepartureCandidateAudit_20261125'
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native['passed']==1 and native['failed']==0 and not native['timeout600']
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledNewtonSolver.cls',
             'src/Solver/CCoupledDeformationContinuation.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
assert (root/'Input/ExistingGuideCandidate.bin').read_bytes()==(parent/'Native/Departure.Candidate1.bin').read_bytes()
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
guide_raw=(root/'Native/Departure.Candidate1.bin').read_bytes();assert guide_raw[:16]==b'NDMDEPARTCAND001'
reader=Reader(guide_raw);guide_meta=reader.VectorTest();accepted=reader.VectorTest();guide=reader.VectorTest()
projection=reader.VectorTest();old_projection=reader.VectorTest();assert reader.position==len(guide_raw)
raw=(root/'Native/GuideReturn.Result.bin').read_bytes();assert raw[:16]==b'NDMGUIDERET00001'
reader=Reader(raw);summary=reader.VectorTest();before=reader.VectorTest();after=reader.VectorTest()
field=reader.VectorTest();actual=reader.VectorTest();assert reader.position==len(raw) and len(summary)==11
spent,available,used,solved,factor,coordinate,target,revision_before,revision_after,guide_factor,seconds=summary
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
assert len(original['fields'])==46 and np.array_equal(accepted,original['fields'][-1])
assert guide_factor==guide_meta[0] and target==guide_meta[6] and guide_meta[5]<=guide_meta[4]
assert spent==before[0] and used==after[0] and available==original['options'][2]-spent
assert 0<=used<=available and spent+used<=30 and revision_before==revision_after and solved in (0,1)
proof=dict(native=native,source_states=46,primary_steps_spent_on_auxiliary=int(spent),remaining_primary_steps=int(available),
    primary_steps_used_on_return=int(used),total_primary_steps=int(spent+used),primary_limit=30,
    auxiliary_effort=before.tolist(),return_effort=after.tolist(),return_seconds=float(seconds),
    original_target=float(target),solved_original_target=bool(solved),returned_lambda=float(factor),
    actual_control=float(coordinate),auxiliary_lambda=float(guide_factor),accepted_revision_unchanged=int(revision_before),
    repeated_guide_candidate_bitidentical_to749=guide_raw==(root/'Input/ExistingGuideCandidate.bin').read_bytes(),
    source_result_sha256=sha256(raw).hexdigest().upper(),numerical_kernel_exact749=True,
    working_lambda_guard_not_changed=True,auxiliary_state_not_admitted_as_history=True,
    no_cold_capacity_search=True,physical_capacity_certified=False,whole_goal_completed=False)
if solved:
    mod=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=mod['TestValidate'](root,'GuideReturn.Sequence')
    assert len(fields)==47 and mechanics['details'][0]=='DiagnosticReSolve'
    np.testing.assert_allclose(fields[:-1],original['fields'],rtol=1e-14,atol=1e-300)
    np.testing.assert_allclose(fields[-1],field,rtol=1e-14,atol=1e-300)
    bound=next(float(r[3]) for r in mod['TestRows'](root/'Native/GuideReturn.Sequence.csv') if r[0]=='R' and r[1]=='47')
    control=float(old_projection@field);assert abs(control-coordinate)<=64*np.finfo(float).eps*max(abs(control),abs(coordinate))
    assert abs(control-target)*float(original['calibration'][0])<=bound+1e-9
    proof.update(independent_NDS_and_original_control_verified=True,independent_return_equilibrium=mechanics['states'][-1],
        control_residual_MPa=float(abs(control-target)*original['calibration'][0]),control_bound_MPa=bound,
        signed_advance_in_original_projection=float(old_projection@(field-accepted)),
        auxiliary_guide_excluded_from47_state_return_sequence=True)
else:
    assert len(field)==len(actual)==1 and not (root/'Native/GuideReturn.Sequence.csv').exists()
    proof.update(independent_NDS_and_original_control_verified=False,no_returned_NDS=True,
        native_return_details=(root/'Native/GuideReturn.Outcome.txt').read_text('cp1251'))
(root/'DepartureGuideReturnProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('750 RETURN',bool(solved),'PRIMARY STEPS',int(spent),'+',int(used),'<=30; LAMBDA',factor,'NO PHYSICAL LIMIT CLAIM')
