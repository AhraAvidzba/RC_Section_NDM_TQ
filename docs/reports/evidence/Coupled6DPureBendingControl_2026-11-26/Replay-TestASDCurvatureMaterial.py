"""ДЛЯ ТЕСТОВ. Сверяет все принятые тензоры кривой изгиба с настоящим ASDConcrete3D OpenSees без решения равновесия."""
from pathlib import Path
from argparse import ArgumentParser
from hashlib import sha256
from time import perf_counter
import json
import openseespy.opensees as ops
import openseespywin.opensees as native

p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True);p.add_argument('--post-curve-force',action='store_true');args=p.parse_args();root=args.root
input_name='PostCurveForceMaterialReplayInput.json' if args.post_curve_force else 'CurvatureMaterialReplayInput.json'
output_name='OpenSeesPostCurveForceMaterialReplay.json' if args.post_curve_force else 'OpenSeesCurvatureMaterialReplay.json'
raw=(root/'Native'/input_name).read_bytes();wire=json.loads(raw);contract=wire['material_contract']
assert contract['integration_mode']=='native_single_update_test' and ops.version()=='3.8.0' and ops.NDTest is native.NDTest
assert wire['native_curve_sha256']==sha256((root/'Native/PureBendingCurvature.bin').read_bytes()).hexdigest().upper()
if args.post_curve_force:assert wire['post_force_binary_sha256']==sha256((root/'Native/PostCurveForce.bin').read_bytes()).hexdigest().upper()
command=['ASDConcrete3D',contract['concrete_E_MPa'],contract['poisson']]
for name,rows in contract['uniaxial_tables'].items():
    for j,field in enumerate(('e','s','d')):command.extend(['-'+name+field,*(row[j] for row in rows)])
command+=['-tangent'];ops.wipe();ops.model('basic','-ndm',3,'-ndf',3)
count=len(wire['frames'][0]['strain'])
for tag in range(1,count+1):ops.nDMaterial(command[0],tag,*command[1:])
start=perf_counter();maximum=0.;where=None;observed=[];comparisons=0
for index,frame in enumerate(wire['frames']):
    assert len(frame['strain'])==count
    stresses=[]
    for point,strain in enumerate(frame['strain']):
        xx,yy,zz,xy,xz,yz=strain;tag=point+1
        ops.NDTest('SetStrain',tag,xx,yy,zz,xy,yz,xz)
        assert tuple(ops.NDTest('GetStrain',tag))==(xx,yy,zz,xy,yz,xz)
        xx,yy,zz,xy,yz,xz=ops.NDTest('GetStress',tag);stress=[xx,yy,zz,xy,xz,yz]
        if frame['expected'] is not None:
            error=max(abs(a-b) for a,b in zip(stress,frame['expected'][point]));comparisons+=6
            if error>maximum:maximum=error;where=dict(frame=index+1,point=point+1,expected=frame['expected'][point],observed=stress)
        stresses.append(stress);ops.NDTest('CommitState',tag)
    observed.append(stresses)
elapsed=perf_counter()-start;ops.wipe()
proof=dict(engine_version='3.8.0',installed_binary_sha256=sha256(Path(native.__file__).read_bytes()).hexdigest().upper(),
    replay_script_sha256=sha256(Path(__file__).read_bytes()).hexdigest().upper(),input_sha256=sha256(raw).hexdigest().upper(),
    observed_OpenSees_concrete_stresses=observed,scalar_comparisons=comparisons,maximum_absolute_stress_error_MPa=maximum,
    largest_difference=where,material_tensor_match=maximum<1e-7,comparison_tolerance_MPa=1e-7,elapsed_seconds=elapsed,
    section_equilibrium_solved_by_OpenSees=False,physical_capacity_certified=False,whole_goal_completed=False)
(root/'Native'/output_name).write_text(json.dumps(proof,separators=(',',':'))+'\n',encoding='utf-8')
print(json.dumps({k:v for k,v in proof.items() if k!='observed_OpenSees_concrete_stresses'},sort_keys=True))
assert proof['material_tensor_match']
