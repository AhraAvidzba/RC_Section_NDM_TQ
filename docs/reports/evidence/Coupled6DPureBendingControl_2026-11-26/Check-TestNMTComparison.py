"""ДЛЯ ТЕСТОВ. Независимо проверяет объёмный NMT и сравнивает средние деформации одинаковых центральных областей с холодным секционным ASD."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from itertools import product
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
section=base/'Cold1003ASDNMTCold_20261126';volume=base/'Cold1004OpenSeesNMTCore_20261126'
source=json.loads((volume/'Preparation.json').read_bytes());d=json.loads((volume/'Result.json').read_bytes())
assert d['termination'] is not None
assert sha256((volume/'validate_opensees_hardshear.py').read_bytes()).hexdigest().upper()==source['source_sha256']
target=np.array([28000.,510000.,180000.,1500000.,0.,0.]);np.testing.assert_array_equal(d['section_target'],target)
assert source['section_target']==target.tolist() and source['no_moment_gradient_from_Q']
assert np.all(np.diff([s['factor'] for s in d['states']])>0) and len(d['states'])>1
max_error=0.;tensors=0;first_material_error=0.;E=30000.;nu=.2;G=E/(2*(1+nu));la=E*nu/((1+nu)*(1-2*nu));D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=la
material=json.loads((section/'Input/MaterialContract.json').read_bytes());alpha=4/33

def TestASDFirstStress(strain):
    """Учитывает ненулевые нижние уровни интерполяции первого спектрального шага; его отклик не равен буквально D*epsilon."""
    trial=D@strain;xx,yy,zz,xy,xz,yz=trial;t=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]])
    eig,axes=np.linalg.eigh(t);trace=sum(eig);q=np.sqrt(1.5*sum((eig-trace/3)**2));negative=np.minimum(eig,0)
    ratio=max(5,max(r[1] for r in material['uniaxial_tables']['C'])/max(r[1] for r in material['uniaxial_tables']['T']))
    beta=(1-alpha)*ratio-(1+alpha);nt=sum(negative);nq=np.sqrt(1.5*sum((negative-nt/3)**2))
    floors=material['prepared_stress_floors_MPa']
    excursion=[(alpha*trace+q+beta*eig[-1])/((1-alpha)*E*ratio) if eig[-1]>=floors['T'] else 0.,
               (alpha*nt+nq+3*negative[-1])/((1-alpha)*E) if eig[0]<0 else 0.]
    scales=[]
    for kind,e in zip(('T','C'),excursion):
        rows=np.array(material['uniaxial_tables'][kind]);x=rows[:,0];nom=rows[:,1];tol=min(np.diff(x))*1e-6
        prior_plastic=-floors[kind]/E;current=max(0,e+prior_plastic)
        i=max(1,min(int(np.searchsorted(x+tol,current,side='left')),len(x)-1));fraction=(current-x[i-1])/(x[i]-x[i-1])
        sigma=max(floors[kind],nom[i-1]+fraction*(nom[i]-nom[i-1]));scales.append(sigma/(E*(current-prior_plastic)))
    tensor=(axes*(eig*np.where(eig>0,scales[0],scales[1])))@axes.T
    return np.array([tensor[0,0],tensor[1,1],tensor[2,2],tensor[0,1],tensor[0,2],tensor[1,2]])
for state in d['states']:
    errors=[float(max(abs((np.array(state[key])-state['factor']*target)/[1,180,180,180,1,1]))) for key in ('base_at_cut','lower_cut','upper_cut')]
    assert max(errors)<=d['cut_check_error_N'] and max(abs(np.array(errors)-state['equilibrium_error']))<1e-8
    max_error=max(max_error,max(errors))
    for block in state['cut_fields']:
        stress=np.array(block['stresses']).reshape(8,6)[:,[0,1,2,3,5,4]]
        strain=np.array(block['strains']).reshape(8,6)[:,[0,1,2,3,5,4]]
        assert np.all(np.isfinite(stress)) and np.all(np.isfinite(strain));tensors+=8
        if state is d['states'][0]:
            error=float(np.max(abs(stress-np.array([TestASDFirstStress(e) for e in strain]))))
            assert error<1e-8
            first_material_error=max(first_material_error,error)
    for kind,bar in state['steel'].items():
        limit=(.025 if bar['strain']>=0 else .018) if kind.startswith('long') else .02
        assert abs(bar['util']-abs(bar['strain'])/limit)<1e-12
    for damage in state['prism_damage_max']:
        if damage['value']>0:assert d['nonlinear_zone_mm'][0]<=damage['z_interval_mm'][0] and damage['z_interval_mm'][1]<=d['nonlinear_zone_mm'][1]
assert max(map(max,d['installed_adjustment']))<1e-7
native=json.loads((section/'ASDHardShearColdProof.json').read_bytes());assert native['all_six_resultants_and56plus4_equations_verified']
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((section/'Native/ASDCold.bin').read_bytes())
options=r.VectorTest();dof,points,concrete=map(int,options[:3]);half=options[3];np.testing.assert_array_equal(r.VectorTest(),target)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](section/'Native/ASDCold.Operators.csv',dof,points)
length=L[0,0]/L[1,1]
fields=[]
while True:
    meta=r.VectorTest()
    if meta[0]==0:break
    if not meta[3]:continue
    field=r.VectorTest();r.VectorTest()
    for _ in range(3*points):
        for _ in range(4):r.VectorTest()
    fields.append((meta[2],field))
assert r.position==len(r.raw)
powers=list(product(range(4),repeat=2))

def TestPolynomial(x,y):
    """Вычисляет полный кубический тензорный базис в нормированных поперечных координатах."""
    return np.array([(x/90.)**i*(y/60.)**j for i,j in powers])

A=np.array([TestPolynomial(x,y) for x,y in coords[:concrete,:2]]);assert np.linalg.matrix_rank(A)==16
comparisons=[];X=np.array([-90.,-76.,-38.,0.,38.,76.,90.]);Y=np.array([-60.,-46.,-30.,0.,30.,46.,60.]);gauss=1/np.sqrt(3.)
for factor,field in fields:
    matches=[s for s in d['states'] if abs(s['factor']-factor)<1e-12]
    if not matches:continue
    state=matches[0];strain=np.einsum('pcd,d->pc',B,field);slope=np.einsum('pcd,d->pc',gradient,field)
    coefficients=np.linalg.lstsq(A,strain[:concrete],rcond=None)[0];gradient_coeff=np.linalg.lstsq(A,slope[:concrete],rcond=None)[0]
    assert np.max(abs(A@coefficients-strain[:concrete]))<1e-14
    assert np.max(abs(A@gradient_coeff-slope[:concrete]))<1e-14
    actual=[];expected=[]
    for block in state['cut_fields']:
        x,y,z=block['center_relative_to_cut_mm'];ix=int(np.searchsorted(X,x)-1);iy=int(np.searchsorted(Y,y)-1)
        hx=(X[ix+1]-X[ix])/2;hy=(Y[iy+1]-Y[iy])/2
        P=np.mean([TestPolynomial(x+sx*gauss*hx,y+sy*gauss*hy) for sx,sy in product((-1,1),repeat=2)],axis=0)
        expected.append(P@(coefficients+z/length*gradient_coeff))
        actual.append(np.mean(np.array(block['strains']).reshape(8,6),axis=0)[[0,1,2,3,5,4]])
    actual=np.array(actual);expected=np.array(expected);difference=expected-actual
    comparisons.append(dict(lambda_value=factor,regions=len(actual),
        mean_strain_field_difference_percent=float(100*np.max(abs(difference))/max(np.max(abs(actual)),np.finfo(float).tiny)),
        component_mean_strain_max_absolute_difference=np.max(abs(difference),axis=0).tolist(),
        component_OpenSees_max_mean_absolute_strain=np.max(abs(actual),axis=0).tolist(),
        component_NDM_max_mean_absolute_strain=np.max(abs(expected),axis=0).tolist()))
assert len(comparisons)>=3
proof=dict(section_case=1003,volume_case=1004,section_target=target.tolist(),volume_solver_seconds=d['elapsed_s'],
    maximum_accepted_lambda_OpenSees=max(s['factor'] for s in d['states']),maximum_accepted_lambda_NDM=native['maximum_accepted_lambda'],
    maximum_independent_resultant_error_N=max_error,checked_volume_tensors=tensors,first_load_independent_material_error_MPa=first_material_error,
    volume_termination=d['termination'],volume_last_cut_damage=d['states'][-1]['adjacent_damage'],
    volume_last_max_damage=d['states'][-1]['prism_damage_max'],strain_comparisons=comparisons,
    nodal_and_gauss_shear_order_explicitly_mapped=True,cubic_NDM_strain_fit_residual_below1e_minus14=True,
    moment_gradient_from_Q_absent=True,auxiliary_elastic_regions_not_homogeneous_beam=True,
    numerical_stop_not_physical_capacity=True,physical_capacity_certified=False,whole_goal_completed=False)
(volume/'NMTComparisonProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('NMT COMPARISON VERIFIED',json.dumps({k:v for k,v in proof.items() if k!='strain_comparisons'},sort_keys=True))
print('LAST COMMON FIELD',json.dumps(comparisons[-1],sort_keys=True))
