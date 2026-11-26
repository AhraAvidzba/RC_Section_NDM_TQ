"""ДЛЯ ТЕСТОВ. Заново собирает нелинейные уравнения и с нуля решает боковой стальной предел без VBA-итераций или их полей."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold885SpatialSteelRawControl_20261126')
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'SpatialSteelRaw')
assert native==dict(passed=14,failed=0,macros=1,timeout600=False,own_process_closed=True)
raw=(root/'Native/SpatialSteelRaw.bin').read_bytes();assert raw[:16]==b'NDMSTEELSIDE0001'
reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'](raw)
meta=reader.VectorTest();assert len(meta)==16
dofs,nc,points=map(int,meta[:3]);E,G,cubic,Es,ft,fc,Et,Ec,et,ec,h,absolute,relative=meta[3:]
assert (E,G,cubic,Es,ft,fc,Et,Ec,et,ec,h,absolute,relative)==(30000,15000,1e10,200000,1,1,199000,199000,.0001,.00013,.001,1e-8,1e-8)
assert points==nc+1 and dofs>4
g,L,W,V,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/SpatialSteelRaw.Operators.csv',dofs,points)
physical=dofs-4;op=np.array([B+side*h*gradient for side in (-1,0,1)])
area=1/L[0,0];scales=np.array([area,1/L[1,1],1/L[2,2],1/L[3,3],area,area])
weights=coords[:,2];xy=coords[:,:2]-np.average(coords[:nc,:2],axis=0,weights=weights[:nc]);x,y=xy.T

def TestMaterial(field):
    """Прямой кубический закон матрицы и монотонная двухлинейная сталь; история VBA не читается."""
    strain=np.einsum('spcd,d->spc',op,field)
    diagonal=np.array([E,E,E,G,G,G]);stress=strain*diagonal
    stress[:,:,2]+=cubic*strain[:,:,2]**3
    tangent=np.zeros((3,points,6,6));indices=np.arange(6);tangent[:,:,indices,indices]=diagonal
    tangent[:,:,2,2]+=3*cubic*strain[:,:,2]**2
    steel=np.einsum('spc,pc->sp',strain[:,nc:],axis[nc:]);yielding=np.where(steel>=0,ft,fc)
    slope=np.where(steel>=0,Et,Ec);elastic=abs(steel)<=yielding/Es
    bar=np.where(elastic,Es*steel,np.sign(steel)*(yielding+slope*(abs(steel)-yielding/Es)))
    bar_tangent=np.where(elastic,Es,slope)
    stress[:,nc:]=-stress[:,nc:]+bar[:,:,None]*axis[None,nc:]
    tangent[:,nc:]=-tangent[:,nc:]+bar_tangent[:,:,None,None]*np.einsum('pi,pj->pij',axis[nc:],axis[nc:])[None]
    utilization=abs(steel)/np.where(steel>=0,et,ec)
    return strain,stress,tangent,utilization

def TestEquations(field):
    """Собирает все физические строки и четыре связи из геометрических операторов и независимо вычисленных напряжений."""
    strain,stress,D,util=TestMaterial(field)
    residual=g@field+np.einsum('pcd,pc->d',W,stress[1])+np.einsum('pcd,pc->d',V,(stress[2]-stress[0])/(2*h))
    tangent=g+np.einsum('pcd,pcj,pjk->dk',W,D[1],B)
    tangent+=np.einsum('pcd,pcj,pjk->dk',V,D[2],op[2])/(2*h)
    tangent-=np.einsum('pcd,pcj,pjk->dk',V,D[0],op[0])/(2*h)
    center=stress[1]
    actual=np.array([weights@center[:,2],(weights*y)@center[:,2],(weights*x)@center[:,2],weights@(x*center[:,5]-y*center[:,4]),weights@center[:,4],weights@center[:,5]])
    return residual,tangent,actual,util

rows=[]
while reader.position<len(raw):
    header,field,actual=reader.VectorTest(),reader.VectorTest(),reader.VectorTest()
    assert len(header)==6 and field.shape==(dofs,) and actual.shape==(6,)
    sign,kind,factor,ratio,central,side=header;assert sign in (-1,1) and kind in (0,1,2)
    target=np.array([0,0,0,0,sign*factor*3e7,0]);external=L@target
    residual,K,reconstructed,util=TestEquations(field);residual-=external
    weak=max(np.max(abs(residual[:physical])),E*np.max(abs(residual[physical:])))
    force=float(np.max(abs((reconstructed-target)/scales)));bound=absolute+relative*np.max(abs(external[:physical]))
    assert max(weak,force)<=bound+2e-10,(header,weak,force,bound)
    assert np.max(abs((reconstructed-actual)/scales))<=2e-10
    assert abs(util.max()-ratio)<2e-12 and abs(util[1].max()-central)<2e-12
    assert int(np.argmax(util)//(points-nc))-1==side
    if kind==1:assert util.max()<=1 and util.max()>.99
    if kind==2:assert util.max()>1
    rows.append(dict(sign=int(sign),kind=int(kind),factor=float(factor),side=int(side),utilization=float(util.max()),
                     residual_MPa=float(weak),force_residual_MPa=float(force),bound_MPa=float(bound)))
assert len(rows)==6 and [(r['sign'],r['kind']) for r in rows]==[(-1,0),(-1,1),(-1,2),(1,0),(1,1),(1,2)]
roots=[]
for sign in (-1,1):
    seed,low,high=[r for r in rows if r['sign']==sign];side=seed['side']
    control=axis[nc]@op[side+1,nc];direction=np.array([0,0,0,0,sign*3e7,0]);external=L@direction
    u=np.zeros(dofs);factor=0.;log=[]
    for iteration in range(15):
        residual,K,actual,util=TestEquations(u);residual-=factor*external
        equation=np.r_[residual,E*(control@u-et)];equation[physical:dofs]*=E
        norm=float(np.max(abs(equation)));log.append(norm)
        if norm<1e-10:break
        J=np.zeros((dofs+1,dofs+1));J[:dofs,:dofs]=K;J[:dofs,-1]=-external
        J[physical:dofs]*=E;J[-1,:dofs]=E*control
        correction=np.linalg.solve(J,-equation);u+=correction[:dofs];factor+=correction[-1]
    else:raise AssertionError(('independent boundary solve did not converge',sign,log))
    assert low['factor']<=factor<=high['factor'],(sign,factor,low,high)
    assert abs(util.max()-1)<1e-10
    roots.append(dict(sign=sign,factor=float(factor),iterations=iteration,cold_zero_start=True,
        residual_MPa=norm,iteration_residuals=log,native_bracket=[low['factor'],high['factor']],
        governing_side=side,maximum_steel_utilization=float(util.max()),field=u.tolist()))
proof=dict(native=native,raw_sha256=sha256(raw).hexdigest().upper(),raw_Double_fields=True,raw_states=rows,
    matrix_law_is_synthetic_cubic=True,all_six_forces_and_all_physical_and_gauge_equations=True,
    material_response_reimplemented_independently=True,independent_zero_start_boundary_roots=roots,
    no_VBA_fields_used_to_initialize_independent_roots=True,physical_HardShear_capacity_certified=False,no_release_promotion=True)
(root/'SpatialSteelRawControlProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('SIX RAW STATES VERIFIED;INDEPENDENT ZERO-START SIDE STEEL LIMITS',[(r['sign'],r['factor'],r['iterations']) for r in roots])
