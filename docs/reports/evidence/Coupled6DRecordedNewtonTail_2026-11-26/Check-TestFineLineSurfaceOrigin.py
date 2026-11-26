"""ДЛЯ ТЕСТОВ. Проверяет поверхность точки5 на положительной станции в действительном новом отказе929, без нового вызова материалов."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from argparse import ArgumentParser
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--output',type=Path,required=True);output=p.parse_args().output
scan=base/'Cold926RecordedTailStressScan_20261126';tail=base/'Cold929RecordedTailFineLine_20261126'
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](scan/'Input/Reverse1.bin')
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](scan/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+t*.001*gradient for t in (-1,0,1)])
r=R((scan/'Native/ArcScan.States.bin').read_bytes());states={}
assert r.raw[:16]==b'NDMARCSCAN000001'
for _ in range(129):
    point,side=r.VectorTest();states[int(point),int(side)]=r.VectorTest()
assert r.position==len(r.raw)
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu))
D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
code=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig')
begin=code.index('def TestYield(');end=code.index('def TestAssemble(',begin)
exec(compile(code[begin:end],str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())
r=R((tail/'Native/NewtonRejected.00001.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
np.testing.assert_array_equal(accepted,source['fields'][-1])
rows=[]
for side in (-1,0,1):
    rows.append(dict(point=5,station_z_over_L=side*.001,side=side,
        current_trial_yield_MPa=TestYield(5,side,current),
        minimum_fraction_trial_yield_MPa=TestYield(5,side,current+delta[:60]/65536)))
assert rows[2]['current_trial_yield_MPa']<0<rows[2]['minimum_fraction_trial_yield_MPa']
lo=0.;hi=1/65536
for _ in range(60):
    mid=(lo+hi)/2
    if mid==lo or mid==hi:break
    if TestYield(5,1,current+mid*delta[:60])>0:hi=mid
    else:lo=mid
files=[tail/'Native/NewtonRejected.00001.bin',scan/'Native/ArcScan.States.bin',scan/'Native/EnvelopeEnergy.Accepted.Operators.csv',scan/'Input/Reverse1.bin']
proof=dict(recorded_source='929/Native/NewtonRejected.00001.bin',accepted25_field_bitexact=True,
    surfaces=rows,plus_station_crossing_fraction_bracket=[lo,hi],minimum_fraction=1/65536,
    current_lambda=float(meta[2]),initial_physical_norm_MPa=float(max(abs(residual[:56]))),
    only_trial_surface_not_actual_stress_continuity_checked=True,
    no_new_material_evaluation_or_NDS_search=True,physical_capacity_certified=False,
    source_hashes={x.relative_to(base).as_posix():sha256(x.read_bytes()).hexdigest().upper() for x in files})
output.write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('NEW ACTUAL STOP CROSSES POINT5 PLUS SURFACE EVEN AT MINIMUM FRACTION;BRACKET',lo,hi)
