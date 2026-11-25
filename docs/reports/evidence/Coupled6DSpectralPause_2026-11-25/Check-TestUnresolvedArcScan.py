"""ДЛЯ ТЕСТОВ. Независимо проверяет поверхности и упругий спектральный переход в отклонённых дуговых полях состояния86."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold834UnresolvedArcScan_20261125';parent=base/'Cold833ContinuedOutgoingWalk_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledDeformationContinuation.cls',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];source=read(root/'Input/Reverse1.bin')
assert len(source['fields'])==86
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=Reader((root/'Input/ArcScanFields.bin').read_bytes());r.position=0
count=r.VectorTest();assert count[0]==6
cases=[r.VectorTest() for _ in range(6)];assert r.position==len(r.raw)
r=Reader((root/'Native/ArcScan.States.bin').read_bytes());assert r.raw[:16]==b'NDMARCSCAN000001'
parameters={}
for _ in range(129):
    point,side=r.VectorTest();state=r.VectorTest();assert len(state)==10
    parameters[int(point),int(side)]=state
assert r.position==len(r.raw)
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
g,L,w,v,coords,axis,b,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
E,nu=source['calibration'][:2];off=E*nu/((1+nu)*(1-2*nu));G=E/(2*(1+nu))
D=np.diag([off+2*G]*3+[G]*3);D[:3,:3]+=off*(np.ones((3,3))-np.eye(3))
levels=(0.,.03125,.0625,.125,.25,.5,.75,1.)
rows=validate['TestRows'](root/'Native/ArcScan.csv');assert len(rows)==6*8*129*2
yields={};stress={};maximum_yield_difference=0.

def TestTensor(values):
    """Разворачивает инженерный вектор напряжений в симметричный тензор без изменения единиц МПа."""
    xx,yy,zz,xy,xz,yz=values
    return np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]])

def TestTrial(case,level,point,side):
    """Вычисляет пробный упругий тензор и поверхность от точной принятой пластической истории."""
    field=source['fields'][-1]+level*(cases[case-1]-source['fields'][-1])
    alpha,gamma,*_=parameters[point,side];plastic=parameters[point,side][2:8];cc,ct=parameters[point,side][8:]
    sigma=D@((b[point-1]+side*.001*gradient[point-1])@field-plastic)
    tensor=TestTensor(sigma);tr=np.trace(tensor);ev=np.linalg.eigvalsh(tensor)
    q=np.sqrt(1.5*np.sum((tensor-np.eye(3)*tr/3)**2));high=ev[-1]
    term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
    value=float((alpha*tr+q+term)/(1-alpha)-cc)
    scale=(abs(alpha*tr)+q+abs(term))/(1-alpha)+abs(cc)
    return value,sigma,ev,max(2e-12,256*np.finfo(float).eps*scale)

for row in rows:
    kind,case,level,point,side=row[:5];case,point,side=int(case),int(point),int(side);level=float(level)
    assert level in levels
    key=(case,level,point,side)
    if kind=='Y':
        assert row[5]=='True';value=float(row[6]);independent,_,_,bound=TestTrial(case,level,point,side)
        maximum_yield_difference=max(maximum_yield_difference,abs(independent-value));assert abs(independent-value)<=bound
        yields[key]=value
    elif kind=='S':
        assert int(row[5])==0;stress[key]=np.array(row[6:12],float)
    else:raise AssertionError(kind)
reload=[];ordinary=[];spectral=[]
for case in range(1,7):
    for point in range(1,44):
        for side in (-1,0,1):
            phi=np.array([yields[case,x,point,side] for x in levels])
            if abs(phi[0])<=32e-12 and min(phi[1:-1])<-32e-12 and phi[-1]>32e-12:reload.append([case,point,side])
            if phi[0]<-32e-12 and phi[-1]>32e-12:ordinary.append([case,point,side])
            eigen=np.array([np.linalg.eigvalsh(TestTensor(stress[case,x,point,side])) for x in levels])
            for principal in np.flatnonzero(eigen[0]*eigen[-1]<0):
                spectral.append(dict(case=case,point=point,side=side,principal_from_high=3-int(principal),
                    native_principal_MPa=eigen[:,principal].tolist(),yield_values_MPa=phi.tolist(),
                    all_samples_elastic=bool(max(phi)<-32e-12)))
assert not reload and not ordinary
assert len(spectral)==18 and all(x['point']==3 and x['principal_from_high']==1 and x['all_samples_elastic'] for x in spectral)
# Матрица S выводит энергетические координаты CDP, а не множители потери жёсткости: их нельзя подставлять вместо d_t/d_c.
# Проверяется положительная скалярная связь фактического напряжения с независимо восстановленным упругим тензором и совпадение знака его старшего значения.
maximum_collinearity_error=0.
for item in spectral:
    case,point,side=item['case'],item['point'],item['side']
    for level in levels:
        _,effective,ev,_=TestTrial(case,level,point,side)
        measured=stress[case,level,point,side]
        factor=float(effective@measured/(effective@effective));assert factor>0
        reconstructed=factor*effective
        error=float(np.max(abs(reconstructed-stress[case,level,point,side])))
        maximum_collinearity_error=max(maximum_collinearity_error,error)
        assert error<=1e-10
        native_high=np.linalg.eigvalsh(TestTensor(measured))[-1]
        assert np.sign(native_high)==np.sign(ev[-1])
proof=dict(native=native,all_working_methods_and_materials_exact833=True,source86_equilibria_independently_rebuilt=mechanics,
    all6192_trial_yield_values_independently_rebuilt=True,maximum_yield_difference_MPa=maximum_yield_difference,
    no_new_yield_surface_crossing_on_any_sampled_segment=True,no_unload_reload_yield_candidate_on_samples=True,
    spectral_sign_changes=spectral,concrete_point3_elastic_largest_principal_crossing_at_all_three_stations=True,
    actual_stress_positive_collinearity_with_independent_elastic_trial_verified=True,
    maximum_collinearity_error_MPa=maximum_collinearity_error,
    CDP_energy_coordinates_not_substituted_for_damage_factors=True,
    sampled_segments_not_accepted_equilibria=True,no_new_global_NDS_or_material_acceptance=True,
    spectral_event_not_yet_supported_by_CDP_arc_fallback=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'UnresolvedArcScanProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('834 INDEPENDENT YIELD6192; NO NEW YIELD CROSSING; POINT3 ELASTIC MAX PRINCIPAL CROSSES ZERO IN ALL THREE STATIONS; COLLINEARITY ERROR',maximum_collinearity_error)
