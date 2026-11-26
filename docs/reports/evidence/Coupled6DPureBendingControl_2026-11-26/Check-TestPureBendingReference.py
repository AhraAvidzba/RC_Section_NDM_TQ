"""ДЛЯ ТЕСТОВ. Независимо интегрирует плоское предельное НДС и находит тот же физический предел вложенными скалярными уравнениями."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
from scipy.optimize import brentq

here=Path(__file__).resolve().parent;root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1005PureBendingReference_20261126')
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
assert 'PLANAR_REFERENCE:' in text
status=(root/'Native/PlanarReference.txt').read_text('cp1251')
assert 'CONVERGED;True;UltimateStrain;' in status
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Native/PlanarReference.bin').read_bytes())
assert r.raw[:16]==b'NDMPLANEMOM00001'
meta,field,limits=r.VectorTest(),r.VectorTest(),r.VectorTest()
lam,moment,N,My,seconds=meta;eps,kx,ky=field[:3];nc,ns,npc,nps=map(int,limits[4:8])
concrete=np.array([r.VectorTest() for _ in range(npc)]);steel=np.array([r.VectorTest() for _ in range(nps)])
geometry=np.array([r.VectorTest() for _ in range(nc)]);bars=np.array([r.VectorTest() for _ in range(ns)])
assert r.position==len(r.raw) and nc==400 and ns==2 and moment>0 and lam>0
assert abs(moment-lam*1e6)<limits[9]
assert np.max(concrete[:,0])==0 and concrete[-1,1]==0

def TestPlane(epsilon0,curvature,curvature_y=0.):
    """Повторяет интегралы напряжений из фактических узлов диаграмм; бетон в точках стали вычитается отдельно."""
    ec=epsilon0+curvature*geometry[:,1]+curvature_y*geometry[:,0];es=epsilon0+curvature*bars[:,1]+curvature_y*bars[:,0]
    sc=np.interp(ec,concrete[:,0],concrete[:,1]);ss=np.interp(es,steel[:,0],steel[:,1])-np.interp(es,concrete[:,0],concrete[:,1])
    actual=np.array([geometry[:,2]@sc+bars[:,2]@ss,(geometry[:,2]*geometry[:,1])@sc+(bars[:,2]*bars[:,1])@ss,
                     (geometry[:,2]*geometry[:,0])@sc+(bars[:,2]*bars[:,0])@ss])
    usage=max(max(-ec)/abs(limits[0]),max(-es)/abs(limits[2]),max(es)/limits[3])
    return actual,usage,ec,es

actual,usage,ec,es=TestPlane(eps,kx,ky)
assert abs(ky)<1e-12
assert np.max(abs(actual-field[3:6])/limits[8:11])<=1.000001
assert abs(actual[0]-N)<=limits[8] and abs(actual[1]-moment)<=limits[9] and abs(actual[2]-My)<=limits[10]
np.testing.assert_allclose([ec.min(),ec.max(),es.min(),es.max()],field[6:10],atol=1e-13,rtol=0)
assert min(abs(ec.min()-limits[0]),abs(es.min()-limits[2]),abs(es.max()-limits[3]))<=1.1e-7

def TestAxialBalance(curvature):
    """Ищет единственную осевую деформацию при заданной кривизне без использования готового поля VBA."""
    span=.1+abs(curvature)*120
    return brentq(lambda epsilon0:TestPlane(epsilon0,curvature)[0][0],-span,span,xtol=1e-15,rtol=1e-14)

independent_k=brentq(lambda curvature:TestPlane(TestAxialBalance(curvature),curvature)[1]-1.,0.,.01,xtol=1e-15,rtol=1e-13)
independent_eps=TestAxialBalance(independent_k);independent,independent_usage,_,_=TestPlane(independent_eps,independent_k)
relative_error=abs(independent[1]-moment)/moment;assert relative_error<.001
proof=dict(native=native,reference_method='stock CCapacitySolver UltimateStrain',reference_limit_state=status.splitlines()[0],
    Mx_ultimate_Nmm=moment,Mx_ultimate_kNm=moment/1e6,lambda_ultimate=lam,solver_seconds=seconds,
    actual_resultants_N_Nmm=actual.tolist(),deformation_plane=[eps,kx,ky],physical_strain_utilization=float(usage),
    concrete_curve=concrete.tolist(),steel_curve=steel.tolist(),concrete_points=nc,longitudinal_bars=ns,
    artificial_diagram_extension=False,tension_concrete_in_planar_strength=False,
    independent_Mx_Nmm=float(independent[1]),independent_deformation_plane=[independent_eps,independent_k,0.],
    independent_usage=float(independent_usage),relative_independent_moment_error=float(relative_error),
    source_unchanged=True,planar_reference_physical_limit_verified=True,coupled_capacity_not_certified=True,whole_goal_completed=False)
(root/'PureBendingReferenceProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PLANAR PURE BENDING VERIFIED',json.dumps({k:v for k,v in proof.items() if k not in ('concrete_curve','steel_curve')},sort_keys=True))
