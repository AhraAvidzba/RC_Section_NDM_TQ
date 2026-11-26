"""ДЛЯ ТЕСТОВ. Проверяет пересечения всех пробных поверхностей в новом фактическом отказе934, без повторного вычисления материала."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from argparse import ArgumentParser
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--output',type=Path,required=True);output=p.parse_args().output
scan=base/'Cold930FineLinePlusSurface_20261126';tail=base/'Cold934ConsistentSpatialLine_20261126'
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](scan/'Input/Reverse1.bin')
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](scan/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+t*.001*gradient for t in (-1,0,1)])
r=R((scan/'Native/ArcScan.States.bin').read_bytes());states={}
for _ in range(129):
    point,side=r.VectorTest();states[int(point),int(side)]=r.VectorTest()
assert r.position==len(r.raw)
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu))
D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
code=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig')
begin=code.index('def TestYield(');end=code.index('def TestAssemble(',begin);exec(code[begin:end],globals())
path=tail/'Native/NewtonRejected.00003.bin';r=R(path.read_bytes());r.position=0
meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
A=np.array([r.VectorTest() for _ in range(61)]);end=r.VectorTest();count=int(r.VectorTest()[0])
probes=[[r.VectorTest() for _ in range(3)] for _ in range(count)];assert r.position==len(r.raw)
np.testing.assert_array_equal(accepted,source['fields'][-1]);info,trial,actual=probes[-1]
assert info[0]==end[-1]==1/65536
crossings=[];near=[]
for point in range(1,44):
    for side in (-1,0,1):
        before=TestYield(point,side,current);after=TestYield(point,side,trial)
        item=dict(point=point,station_z_over_L=side*.001,side=side,current_trial_yield_MPa=before,trial_yield_MPa=after)
        if before*after<0:crossings.append(item)
        near.append(item)
near.sort(key=lambda x:abs(x['current_trial_yield_MPa']))
proof=dict(source_record='934/Native/NewtonRejected.00003.bin',accepted25_history_field_exact=True,
    current_lambda=float(meta[2]),governing_row=int(np.argmax(abs(residual[:56])))+1,
    initial_norm_MPa=float(end[1]),minimum_fraction=float(end[-1]),last_valid_norm_MPa=float(info[2]),
    near_trial_surfaces=near[:12],surface_crossings_at_minimum_fraction=crossings,
    no_crossing_not_proof_of_material_smoothness=True,no_new_material_evaluation_or_NDS=True,
    physical_capacity_certified=False,
    source_hashes={x.relative_to(base).as_posix():sha256(x.read_bytes()).hexdigest().upper() for x in
        (path,scan/'Native/ArcScan.States.bin',scan/'Native/EnvelopeEnergy.Accepted.Operators.csv',scan/'Input/Reverse1.bin')})
output.write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('NEXT ACTUAL FAILURE;R',end[1],'CROSSINGS',crossings,'NEAR',near[:3])
