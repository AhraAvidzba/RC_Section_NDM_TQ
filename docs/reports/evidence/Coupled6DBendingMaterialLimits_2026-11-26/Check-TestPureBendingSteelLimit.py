"""ДЛЯ ТЕСТОВ. Подтверждает полное пространственное равновесие и фактическую предельную деформацию выбранного продольного стержня."""
from pathlib import Path
from hashlib import sha256
import runpy,json,sys,numpy as np
from argparse import ArgumentParser

here=Path(__file__).resolve().parent;p=ArgumentParser(description=__doc__)
p.add_argument('--root',type=Path,default=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1016PureBendingSteelLimitFixed_20261126'));root=p.parse_args().root
sys.argv=['Check-TestPureBendingCurvature.py','--root',str(root)]
scope=runpy.run_path(str(here/'Check-TestPureBendingCurvature.py'))
r=scope['R']((root/'Native/PureBendingSteelLimit.bin').read_bytes());assert r.raw[:16]==b'NDMBENDSTEEL0001'
mu,point,side,target,prior_strain,prior_factor=r.VectorTest();point=int(point);side=int(side)
assert mu==scope['mu'] and prior_factor==scope['records'][-1]['lambda_value'] and 25<point<=43 and side in (-1,0,1)
assert scope['axis'][point-1].tolist()==[0.,0.,1.,0.,0.,0.]
projection=r.VectorTest();np.testing.assert_allclose(projection,scope['op'][side+1,point-1,2],atol=1e-15,rtol=0)
assert abs(projection@scope['history'][-1]-prior_strain)<1e-15
solved,factor,iterations,evaluations,norm=r.VectorTest();assert solved==1
state=scope['TestReadState'](r);assert r.position==len(r.raw)
scope['history'].append(state[0]);scope['TestMechanics'](state,factor)
record=scope['records'][-1];control=30000*abs(projection@state[0]-target)
assert control<=record['original_bound_MPa'] and abs(record['steel_strain_utilization']-1)<1e-6
assert record['longitudinal_concrete_strain_ratio_to_planar_limit']<1
assert target==.025 and abs(state[2][side+1,point-1,2]-.025)<1e-10
proof=dict(native=scope['native'],reference_Mx_kNm=mu/1e6,limit_Mx_kNm=record['Mx_kNm'],limit_ratio=factor,
    critical_longitudinal_point=point,critical_station=side,target_strain=target,actual_strain=float(projection@state[0]),
    control_residual_MPa=control,iterations=int(iterations),evaluations=int(evaluations),reported_norm_MPa=norm,
    independently_verified_equilibrium=record,longitudinal_strain_limit_equilibrium_confirmed=True,
    no_future_equilibrium_input=True,source_unchanged=True,physical_capacity_of_all6_not_certified=True,
    material_laws_differ_from_planar_reference=True,global_peak_not_certified=True,whole_goal_completed=False)
(root/'PureBendingSteelLimitProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
scope['material_frames'].append(dict(strain=state[2].reshape(-1,6).tolist(),expected=state[4].reshape(-1,6).tolist()))
wire=json.loads((root/'Native/CurvatureMaterialReplayInput.json').read_bytes());wire['frames']=scope['material_frames']
wire['steel_limit_binary_sha256']=sha256(r.raw).hexdigest().upper()
(root/'Native/SteelLimitMaterialReplayInput.json').write_text(json.dumps(wire,separators=(',',':'))+'\n',encoding='utf-8')
print('ASD STEEL LIMIT CONFIRMED',json.dumps(proof,sort_keys=True))
