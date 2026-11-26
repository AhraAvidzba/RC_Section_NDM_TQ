"""ДЛЯ ТЕСТОВ. Передаёт библиотеке OpenSees только таблицы и принятые истории деформаций холодного ASD, сравнивая все бетонные тензоры."""
from pathlib import Path
from argparse import ArgumentParser
from hashlib import sha256
from time import perf_counter
import json,struct
import openseespy.opensees as ops
import openseespywin.opensees as native

p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True);root=p.parse_args().root
contract=json.loads((root/'Input/MaterialContract.json').read_bytes())
assert contract['integration_mode']=='native_single_update_test'
binary=(root/'Native/ASDCold.bin').read_bytes();assert binary[:16]==b'NDMASDCOLD000001'
position=16

def TestVector():
    """Читает точные Double, отклоняя незаконченный протокол нативного расчёта."""
    global position
    count,=struct.unpack_from('<i',binary,position);position+=4;assert 0<count<=10000
    values=struct.unpack_from('<'+'d'*count,binary,position);position+=8*count
    return values

options=TestVector();points=int(options[1]);target=TestVector();frames=[];expected=[];factors=[]
while True:
    meta=TestVector()
    if meta[0]==0:break
    if not meta[3]:continue
    field=TestVector();actual=TestVector();strains=[];stresses=[]
    for k in range(3*points):
        point,side,code=TestVector();assert point==k%points+1 and side==k//points-1 and code==0
        strains.append(TestVector());TestVector();stresses.append(TestVector())
    frames.append(strains);expected.append(stresses);factors.append(meta[2])
assert position==len(binary) and len(frames)>1
command=['ASDConcrete3D',contract['concrete_E_MPa'],contract['poisson']]
for name,rows in contract['uniaxial_tables'].items():
    for j,field in enumerate(('e','s','d')):command.extend(['-'+name+field,*(row[j] for row in rows)])
command+=['-tangent']
assert ops.version()=='3.8.0' and ops.NDTest is native.NDTest
ops.wipe();ops.model('basic','-ndm',3,'-ndf',3)
for tag in range(1,3*points+1):ops.nDMaterial(command[0],tag,*command[1:])
started=perf_counter();observed=[];maximum=0.;where=None
for frame_index,frame in enumerate(frames):
    observed_frame=[]
    for point_index,strain in enumerate(frame):
        tag=point_index+1;xx,yy,zz,xy,xz,yz=strain
        ops.NDTest('SetStrain',tag,xx,yy,zz,xy,yz,xz)
        assert tuple(ops.NDTest('GetStrain',tag))==(xx,yy,zz,xy,yz,xz)
        xx,yy,zz,xy,yz,xz=ops.NDTest('GetStress',tag);stress=[xx,yy,zz,xy,xz,yz]
        error=max(abs(a-b) for a,b in zip(stress,expected[frame_index][point_index]))
        if error>maximum:
            maximum=error;where=dict(frame=frame_index+1,lambda_value=factors[frame_index],point=point_index%points+1,
                                     side=point_index//points-1,native_VBA_stress=expected[frame_index][point_index],OpenSees_stress=stress)
        observed_frame.append(stress);ops.NDTest('CommitState',tag)
    observed.append(observed_frame)
elapsed=perf_counter()-started;ops.wipe()
result=dict(engine_version='3.8.0',installed_binary_sha256=sha256(Path(native.__file__).read_bytes()).hexdigest().upper(),
    replay_script_sha256=sha256(Path(__file__).read_bytes()).hexdigest().upper(),
    native_binary_sha256=sha256(binary).hexdigest().upper(),material_command=command,factors=factors,
    strain_histories=frames,expected_VBA_concrete_stresses=expected,observed_OpenSees_concrete_stresses=observed,
    scalar_comparisons=len(frames)*3*points*6,maximum_absolute_stress_error_MPa=maximum,largest_difference=where,
    material_tensor_match=maximum<1e-7,comparison_tolerance_MPa=1e-7,elapsed_seconds=elapsed,
    section_equilibrium_solved_by_OpenSees=False,whole_goal_completed=False,physical_capacity_certified=False)
(root/'Native/OpenSeesNativeMaterialReplay.json').write_text(json.dumps(result,separators=(',',':'))+'\n',encoding='utf-8')
print('OPENSEES MATERIAL REPLAY',json.dumps({k:v for k,v in result.items() if k not in ('strain_histories','expected_VBA_concrete_stresses','observed_OpenSees_concrete_stresses','material_command','factors')},sort_keys=True))
if not result['material_tensor_match']:raise SystemExit(1)
