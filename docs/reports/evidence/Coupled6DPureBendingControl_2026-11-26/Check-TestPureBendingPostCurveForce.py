"""ДЛЯ ТЕСТОВ. Проверяет точные силовые пробы после собственной кривой, сохраняя историю стали и все пространственные уравнения."""
from pathlib import Path
from hashlib import sha256
import runpy,json,sys

here=Path(__file__).resolve().parent;root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1013PureBendingPostCurveForce_20261126')
sys.argv=['Check-TestPureBendingCurvature.py','--root',str(root)]
scope=runpy.run_path(str(here/'Check-TestPureBendingCurvature.py'))
r=scope['R']((root/'Native/PostCurveForce.bin').read_bytes());assert r.raw[:16]==b'NDMBENDPOST00001'
mu,initial,count=r.VectorTest();assert mu==scope['mu'] and initial==scope['records'][-1]['lambda_value'] and count==6
cases=[]
for expected in (1.,1.02,1.05,1.1,1.2,1.3):
    ratio,solved,iterations,evaluations,norm=r.VectorTest();assert ratio==expected and solved in (0,1)
    case=dict(ratio=ratio,requested_Mx_kNm=ratio*mu/1e6,equilibrium_at_requested_moment=bool(solved),
        iterations=int(iterations),evaluations=int(evaluations),reported_norm_MPa=norm)
    if solved:
        state=scope['TestReadState'](r);scope['history'].append(state[0]);scope['TestMechanics'](state,ratio)
        scope['material_frames'].append(dict(strain=state[2].reshape(-1,6).tolist(),expected=state[4].reshape(-1,6).tolist()))
        record=scope['records'][-1];case.update(verified_equilibrium=record,
            steel_physical_limit_exceeded=record['steel_strain_utilization']>1,
            reference_concrete_compression_limit_exceeded=record['longitudinal_concrete_strain_ratio_to_planar_limit']>1)
    cases.append(case)
assert r.position==len(r.raw)
proof=dict(native=scope['native'],initial_equilibrium_ratio=initial,cases=cases,
    maximum_requested_ratio_with_equilibrium=max((c['ratio'] for c in cases if c['equilibrium_at_requested_moment']),default=None),
    all_successful_forces_and56plus4_equations_independently_verified=True,
    force_probes_follow_this_run_curvature_history=True,cold_default_load_control_not_equivalent=True,
    post_curve_binary_sha256=sha256(r.raw).hexdigest().upper(),source_unchanged=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'PureBendingPostCurveForceProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
wire=json.loads((root/'Native/CurvatureMaterialReplayInput.json').read_bytes());wire['frames']=scope['material_frames']
wire['post_force_binary_sha256']=proof['post_curve_binary_sha256']
(root/'Native/PostCurveForceMaterialReplayInput.json').write_text(json.dumps(wire,separators=(',',':'))+'\n',encoding='utf-8')
print('POST CURVE FORCE VERIFIED',json.dumps(proof,sort_keys=True))
