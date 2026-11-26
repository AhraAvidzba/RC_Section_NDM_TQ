"""ДЛЯ ТЕСТОВ. Переиспользует неизменяемую книгу931 для касательных точного нового отказа934 у точки11."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import struct,json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold931SurfaceMatrices_20261126';source=base/'Cold934ConsistentSpatialLine_20261126'
scan=base/'Cold930FineLinePlusSurface_20261126';root=base/'Cold935Point11Matrices_20261126'
proofpath=Path.cwd()/'docs/reports/evidence/Coupled6DSurfaceDescent_2026-11-26/RemainingSurfaceOriginProof.json'
proof=json.loads(proofpath.read_bytes())
for name,digest in proof['source_hashes'].items():assert sha256((base/name).read_bytes()).hexdigest().upper()==digest,name
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
path=source/'Native/NewtonRejected.00003.bin';raw=path.read_bytes();r=R(raw);r.position=0
meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
assert meta[7]==1 and meta[8]==1
binary=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](scan/'Input/Reverse1.bin')
np.testing.assert_array_equal(accepted,binary['fields'][-1]);assert len(binary['fields'])==25
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](scan/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+t*.001*gradient for t in (-1,0,1)])
r=R((scan/'Native/ArcScan.States.bin').read_bytes());states={}
for _ in range(129):
    point,side=r.VectorTest();states[int(point),int(side)]=r.VectorTest()
assert r.position==len(r.raw)
E,nu=binary['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu))
D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
code=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig')
begin=code.index('def TestYield(');end=code.index('def TestAssemble(',begin);exec(code[begin:end],globals())
lo=0.;hi=1/65536;assert TestYield(11,-1,current)<0<TestYield(11,-1,current+hi*delta[:60])
for _ in range(60):
    mid=(lo+hi)/2
    if mid==lo or mid==hi:break
    if TestYield(11,-1,current+mid*delta[:60])>0:hi=mid
    else:lo=mid
center=(lo+hi)/2;assert center>1e-6
levels=[0.,center-1e-6,center+1e-6,center+1e-7,center+1e-8,1/65536,1/1024]
positions=[0]
for _ in range(67):
    count,=struct.unpack_from('<i',raw,positions[-1]);positions.append(positions[-1]+4+8*count)
fixture=raw[:positions[-1]]+raw[positions[6]:positions[-1]]
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
values={'Input/FrozenNewton.bin':fixture,'Input/SurfaceMatrixLevels.bin':vector(levels),
 'Input/ActualRejectedNewton.bin':raw,'Input/ActualRefinedNewton.bin':raw,
 'SourceRemainingSurfaceOriginProof.json':proofpath.read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='Exact934Point11IncomingOutgoingMatrices',actual_import_this_pass=['modTestCoupledElasticState'],
 actual_imports_this_replay=[],actual_component_count=257,immutable_compiled_book_reused=True,compiled_source_root=parent.name,
 source_record='934/Native/NewtonRejected.00003.bin',source_matrix_is_refined=True,
 accepted_history25_unchanged=True,all_material_and_numerical_classes_exact931=True,
 point11_minus_surface_crossing_fraction_bracket=[lo,hi],matrix_levels=levels,
 no_new_NDS_search_or_material_acceptance=True,physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'IMMUTABLE931;ACTUAL934 REFINED SYSTEM;POINT11 MINUS CROSSING',lo,hi)
