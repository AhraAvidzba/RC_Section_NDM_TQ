"""ДЛЯ ТЕСТОВ. Независимо собирает локальные вклады фактического прогнозного направления и отделяет бетон замещения стали от реальных бетонных точек."""
from pathlib import Path
from runpy import run_path
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold906Cold32OutgoingStressScan_20261126';parent=base/'Cold900Cold32OutgoingJacobian_20261126'
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
assert 'READONLY_PRIVATE_SCAN_COMPLETED: no new equilibrium search' in text
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
assert (root/'Input/Reverse1.bin').read_bytes()==(parent/'Input/Reverse1.bin').read_bytes()
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes(),str(path)
assert (root/'tests/modTestCoupledElasticState.bas').read_bytes()==(parent/'tests/modTestCoupledElasticState.bas').read_bytes()
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
assert len(fields)==32
assert (root/'Native/EnvelopeEnergy.Accepted.Operators.csv').read_bytes()==(parent/'Native/EnvelopeEnergy.Accepted.Operators.csv').read_bytes()
for name in ('EnvelopeEnergy.Accepted.csv',):
    a=validate['TestRows'](root/'Native'/name);b=validate['TestRows'](parent/'Native'/name)
    assert [r for r in a if r[0] in ('F','S')]==[r for r in b if r[0] in ('F','S')]
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
accepted=source['fields'][-1];factor=source['lambdas'][-1]
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Input/ArcScanFields.bin').read_bytes());r.position=0
assert r.VectorTest().tolist()==[3.];predictions=[r.VectorTest() for _ in range(3)];assert r.position==len(r.raw)
np.testing.assert_array_equal(predictions[0],accepted)
cases=json.loads((root/'Input/ArcScanCases.json').read_bytes());assert [c['case'] for c in cases]==[1,2,3]
levels=(0.,.03125,.0625,.125,.25,.5,.75,1.);stresses={};yields={};failures={}
rows=validate['TestRows'](root/'Native/ArcScan.csv');assert len(rows)==3*8*129*2
for row in rows:
    kind,case,level,point,side=row[:5];key=(int(case),float(level),int(point),int(side))
    assert key[0] in (1,2,3) and key[1] in levels and 1<=key[2]<=43 and key[3] in (-1,0,1)
    if kind=='S':
        assert key not in stresses
        failures[key]=int(row[5]);stresses[key]=np.array(row[6:12],float)
    elif kind=='Y':
        assert key not in yields;yields[key]=(row[5]=='True',float(row[6]))
    else:raise AssertionError(kind)
before=np.zeros((3,43,6));native_bars=np.zeros((3,18))
for row in a:
    if row[0]=='S' and int(row[1])==32:
        side,point=int(row[2])+1,int(row[3])-1;before[side,point]=np.array(row[4:10],float)
        if point>=25:native_bars[side,point-25]=float(row[12])
state_reader=R((root/'Native/ArcScan.States.bin').read_bytes());assert state_reader.raw[:16]==b'NDMARCSCAN000001'
states={}
for _ in range(129):
    point,side=state_reader.VectorTest();state=state_reader.VectorTest();assert len(state)==10;states[int(point),int(side)]=state
assert state_reader.position==len(state_reader.raw)
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu));D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
def TestYield(point,side,field):
    """Воспроизводит только пробную поверхность от принятой пластической истории, без использования нативной функции поверхности."""
    state=states[point,side];alpha,gamma=state[:2];cc,ct=state[8:]
    sigma=D@(op[side+1,point-1]@field-state[2:8]);xx,yy,zz,xy,xz,yz=sigma
    tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);trace=float(np.trace(tensor));high=float(np.linalg.eigvalsh(tensor)[-1])
    q=float(np.sqrt(1.5*np.sum((tensor-np.eye(3)*trace/3)**2)))
    term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
    return float((alpha*trace+q+term)/(1-alpha)-cc)
def TestAssemble(field,lam,stress):
    """Собирает реальные бетонные вклады, вычитание бетона на стальных точках, положительную сталь и остальные уравнения отдельно."""
    strain=np.einsum('spcd,d->spc',op,field);bars=200000*np.einsum('spc,pc->sp',strain[:,25:],axis[25:])
    long=axis[25:,2]==1;fy=np.where(long[None,:],np.where(bars>=0,400,350),240);assert np.all(abs(bars)<fy)
    concrete=np.einsum('pcd,pc->pd',w,stress[1])+np.einsum('pcd,pc->pd',v,(stress[2]-stress[0])/.002)
    steel=bars[:,:,None]*axis[None,25:]
    steel_eq=np.einsum('pcd,pc->pd',w[25:],steel[1])+np.einsum('pcd,pc->pd',v[25:],(steel[2]-steel[0])/.002)
    points=concrete.copy();points[25:]=steel_eq-concrete[25:]
    target=source['offset']+lam*source['direction'];other=g@field-L@target;residual=other+sum(points)
    net=stress[1].copy();net[25:]=steel[1]-stress[1,25:]
    area=coords[:,2];center=np.average(coords[:25,:2],axis=0,weights=area[:25]);x,y=(coords[:,:2]-center).T
    actual=np.array([area@net[:,2],(area*y)@net[:,2],(area*x)@net[:,2],area@(x*net[:,5]-y*net[:,4]),area@net[:,4],area@net[:,5]])
    scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
    return dict(residual=residual,concrete=concrete,steel=steel_eq,points=points,other=other,bars=bars,
                physical_norm=float(max(abs(residual[:56]))),six_force_norm=float(max(abs((actual-target)/scales))))
summary=[];assembled={};yield_error=0.;stress_drift=0.
for case in cases:
    c=case['case'];movement=predictions[c-1]-accepted;items=[]
    for level in levels:
        field=accepted+level*movement;bad=[dict(point=p,side=s,code=failures[c,level,p,s]) for p in range(1,44) for s in (-1,0,1) if failures[c,level,p,s]!=0]
        for point in range(1,44):
            for side in (-1,0,1):
                available,value=yields[c,level,point,side]
                if available:yield_error=max(yield_error,abs(TestYield(point,side,field)-value))
        item=dict(level=level,all_local_returns_available=not bad,failed_points=bad)
        if not bad:
            stress=np.array([[stresses[c,level,p,s] for p in range(1,44)] for s in (-1,0,1)])
            if c==1 or level==0:stress_drift=max(stress_drift,float(np.max(abs(stress-before))))
            value=TestAssemble(field,factor+level*case['lambda_increment'],stress);assembled[c,level]=value
            item.update(physical_residual_MPa=value['physical_norm'],six_force_residual_MPa=value['six_force_norm'])
            if level>0 and (c,0.) in assembled:
                start=assembled[c,0.];delta=(value['residual']-start['residual'])/level
                per_point=(value['points']-start['points'])[:,52]/level;order=np.argsort(abs(per_point))[::-1][:6]
                item.update(row53_derivative_MPa=float(delta[52]),physical_derivative_norm_MPa=float(max(abs(delta[:56]))),
                    row53_real_concrete_MPa=float(sum((value['concrete']-start['concrete'])[:25,52])/level),
                    row53_subtracted_concrete_MPa=float(-sum((value['concrete']-start['concrete'])[25:,52])/level),
                    row53_steel_MPa=float(sum((value['steel']-start['steel'])[:,52])/level),
                    row53_other_MPa=float((value['other'][52]-start['other'][52])/level),
                    largest_row53_points=[dict(point=int(p)+1,kind='real_concrete' if p<25 else 'steel_minus_replaced_concrete',value_MPa=float(per_point[p])) for p in order])
        items.append(item)
    summary.append(dict(case=c,name=case['name'],samples=items))
assert yield_error<2e-10 and stress_drift<1e-11
zero=assembled[1,0.];assert zero['physical_norm']<=mechanics['states'][-1]['bound_MPa']+1e-12
np.testing.assert_allclose(zero['bars'],native_bars,rtol=1e-13,atol=1e-11)
proof=dict(native=native,immutable_actual_compiled_book900_reused=True,actual_components=256,
 source32_history_and_stresses_exact=True,accepted_history_native_guard_unchanged=True,
 independent_yield_max_error_MPa=yield_error,zero_increment_stress_drift_MPa=stress_drift,
 zero_equilibrium_MPa=zero['physical_norm'],segments=summary,all_material_and_kinematic_sources_unchanged=True,
 no_new_NDS_or_material_acceptance=True,directions_are_predictions_not_equilibria=True,
 physical_capacity_certified=False,whole_goal_completed=False)
(root/'Cold32OutgoingStressScanProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('READONLY LOCAL SCAN VERIFIED;YIELD ERROR',yield_error,'STRESS DRIFT',stress_drift)
for segment in summary:
    print(segment['case'],segment['name'])
    for item in segment['samples']:
        print(' LEVEL',item['level'],'AVAILABLE',item['all_local_returns_available'],'R53/STEP',item.get('row53_derivative_MPa'),
              'REAL',item.get('row53_real_concrete_MPa'),'REPLACED',item.get('row53_subtracted_concrete_MPa'),
              'STEEL',item.get('row53_steel_MPa'),'OTHER',item.get('row53_other_MPa'))
