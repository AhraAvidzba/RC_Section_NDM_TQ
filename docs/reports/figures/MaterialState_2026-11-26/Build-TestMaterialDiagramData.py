"""ДЛЯ ТЕСТОВ. Готовит проверенные точки графика из фактических напряжений последнего отказа и всей принятой истории."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold937MaterialDiagramView_20261126';parent=base/'Cold936RecordedYieldCoordinate_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
vld=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,history=vld['TestValidate'](root,'EnvelopeEnergy.Accepted');assert len(history)==25
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+t*.001*gradient for t in (-1,0,1)])
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Input/ArcScanFields.bin').read_bytes());r.position=0
count=int(r.VectorTest()[0]);fields=[r.VectorTest() for _ in range(count)];assert r.position==len(r.raw)
cases=json.loads((root/'Input/ArcScanCases.json').read_bytes());assert len(cases)==count
rows=vld['TestRows'](root/'Native/ArcScan.csv');assert len(rows)==count*8*129*2
stress={};surface={};failures=[]
for row in rows:
    if float(row[2])!=1.:continue
    key=int(row[1]),int(row[3]),int(row[4])
    if row[0]=='S':
        assert key not in stress
        if int(row[5]):failures.append((key,int(row[5])))
        stress[key]=np.array(row[6:12],float)
    elif row[0]=='Y':surface[key]=float(row[6])
    else:raise AssertionError(row[0])
assert not failures
r=R((root/'Input/ActualRejectedNewton.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
A=np.array([r.VectorTest() for _ in range(61)]);end=r.VectorTest();n=int(r.VectorTest()[0])
probes=[[r.VectorTest() for _ in range(3)] for _ in range(n)];assert r.position==len(r.raw)
np.testing.assert_array_equal(current,fields[1]);np.testing.assert_array_equal(accepted,fields[0])
helper=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig');begin=helper.index('def TestAssemble(');finish=helper.index('summary=[];',begin)
exec(compile(helper[begin:finish],str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())
errors=[]
for c,reference in ((2,residual),(count,probes[-1][2])):
    sigma=np.array([[stress[c,p,s] for p in range(1,44)] for s in (-1,0,1)])
    lam=source['lambdas'][-1]+cases[c-1]['lambda_increment']
    rebuilt=TestAssemble(fields[c-1],lam,sigma)
    errors.append(float(max(abs(rebuilt['residual']-reference))))
assert max(errors)<2e-11

def TestTensor(values,engineering_strain=False):
    """Различает инженерные деформации сдвига и тензорные касательные напряжения при вычислении главных значений."""
    a,b,c,d,e,f=values
    if engineering_strain:d/=2;e/=2;f/=2
    return np.array([[a,d,e],[d,b,f],[e,f,c]])

def TestPoint(field,sigma):
    """Получает наибольшие главные значения фактических тензоров; одноосной диаграммой эта проекция не объявляется."""
    strain=op[0,10]@field
    return dict(epsilon1=float(np.linalg.eigvalsh(TestTensor(strain,True))[-1]),
        sigma1_MPa=float(np.linalg.eigvalsh(TestTensor(sigma))[-1]),epsilon_z=float(strain[2]),sigma_z_MPa=float(sigma[2]))

accepted_rows=vld['TestRows'](root/'Native/EnvelopeEnergy.Accepted.csv');accepted_stress={}
for row in accepted_rows:
    if row[0]=='S':accepted_stress[int(row[1]),int(row[2]),int(row[3])]=row
trajectory=[]
for i,field in enumerate(history,1):
    row=accepted_stress[i,-1,11];point=TestPoint(field,np.array(row[4:10],float))
    point.update(state=i,lambda_value=float(source['lambdas'][i-1]),tension_energy_ratio=float(row[10]),compression_energy_ratio=float(row[11]))
    trajectory.append(point)
local=[]
for case,field in zip(cases[1:],fields[1:]):
    point=TestPoint(field,stress[case['case'],11,-1]);point.update(fraction=case['newton_fraction'],trial_yield_MPa=surface[case['case'],11,-1])
    local.append(point)
current_point=local[0];center=json.loads((root/'Preparation.json').read_bytes())['directed_scope']['point11_minus_crossing_fraction']
event=min(local,key=lambda x:abs(x['fraction']-center));assert abs(event['fraction']-center)<1e-18
bars=200000*np.einsum('spc,pc->sp',np.einsum('spcd,d->spc',op,current)[:,25:],axis[25:])
long=axis[25:,2]==1;fy=np.where(long[None,:],np.where(bars>=0,400,350),240);assert np.all(abs(bars)<fy)
steel=[]
for family,mask,label,yield_value in [('longitudinal',long,'Продольная',400.),('transverse',~long,'Поперечная',240.)]:
    p=np.where(mask)[0];sub=bars[:,p];ii,jj=np.unravel_index(np.argmax(abs(sub)),sub.shape);sigma=float(sub[ii,jj])
    steel.append(dict(family=family,label=label,point=int(p[jj])+26,side=int(ii)-1,
        sigma_MPa=sigma,epsilon=sigma/200000,elastic_modulus_MPa=200000.,tension_yield_MPa=yield_value,
        absolute_stress_over_tension_yield=abs(sigma)/yield_value))
data=dict(kind='ActualModelDiagnosticPlot',concrete=dict(point=11,side=-1,station_z_over_L=-.001,
    projection_not_uniaxial_material_law=True,trajectory=trajectory,local=local,current=current_point,event=event),
    steel=steel,source_actual934_record=3,all_steel_on_elastic_branch=True,
    newton_trial_fields_are_not_equilibria=True,physical_capacity_certified=False,
    source_hashes={p.relative_to(base).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in
        (root/'Native/ArcScan.csv',root/'Native/EnvelopeEnergy.Accepted.csv',root/'Native/EnvelopeEnergy.Accepted.Operators.csv',root/'Input/ActualRejectedNewton.bin')})
out=repo/'docs/reports/figures/MaterialState_2026-11-26';out.mkdir(parents=True,exist_ok=True)
(out/'MaterialDiagramData.json').write_text(json.dumps(data,ensure_ascii=False,sort_keys=True,indent=2)+'\n',encoding='utf-8')
proof=dict(native=native,immutable_actual_compiled936_book_reused=True,accepted25_equilibria_reverified=True,
    endpoint_count=count,all_endpoint_material_returns_available=True,source934_base_and_last_probe_reconstruction_errors_MPa=errors,
    principal_tensor_engineering_shear_conversion_explicit=True,original_history_unchanged=True,
    no_new_NDS_or_material_acceptance=True,physical_capacity_certified=False)
(root/'MaterialDiagramViewProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('CURRENT CONCRETE',current_point,'EVENT',event,'STEEL',steel,'RECONSTRUCTION ERRORS',errors)
