"""ДЛЯ ТЕСТОВ. Проверяет масштабирование уточнённых касательных и сравнивает их направления с принятой секущей без решения нового НДС."""
from pathlib import Path
from runpy import run_path
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold864Cold19BranchTangents_20261125';parent=base/'Cold858BalancedInitialStressReplay_20261125';scan=base/'Cold862ColdAcceptedStressContinuity_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and audit['imported_this_pass']==['modTestCoupledElasticState'] and len(audit['other_actual_components_exact_parent_except_identifier_case'])==255
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes(),path
code=(root/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig');old=(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
begin=code.index("' FOR TESTS: baseline entry point only; existing methods retained.\n\n\nPublic Function TestReadColdBranchTangents()")
end=code.index('\nEnd Function\n',begin)+len('\nEnd Function\n');assert code[:begin]+code[end:]==old
for suffix in ('.csv','.Operators.csv'):
    name='Native/EnvelopeEnergy.Accepted'+suffix;assert (root/name).read_bytes()==(scan/name).read_bytes()
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];source=read(root/'Input/Reverse1.bin');assert len(source['fields'])==19
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0
accepted=source['fields'][-1];secant=accepted-source['fields'][-2];length=float(np.sqrt(secant@M@secant))
reference=float(source['lambdas'][-1]);dl=reference-float(source['lambdas'][-2]);assert dl>0 and length>0
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=Reader((scan/'Native/ArcScan.States.bin').read_bytes());assert r.raw[:16]==b'NDMARCSCAN000001'
states={}
for _ in range(129):
    point,side=r.VectorTest();state=r.VectorTest();assert len(state)==10;states[int(point),int(side)]=state
assert r.position==len(r.raw)
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu));D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off

def TestYield(point,side,field):
    """Воспроизводит пробную поверхность от неизменной принятой пластической истории, не обновляя материал."""
    alpha,gamma,*_=states[point,side];plastic=states[point,side][2:8];cc,ct=states[point,side][8:]
    sigma=D@(op[side+1,point-1]@field-plastic);xx,yy,zz,xy,xz,yz=sigma
    tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);trace=float(np.trace(tensor));high=float(np.linalg.eigvalsh(tensor)[-1])
    q=float(np.sqrt(1.5*np.sum((tensor-np.eye(3)*trace/3)**2)))
    term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
    return float((alpha*trace+q+term)/(1-alpha)-cc)

rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows'];yield_error=0
for row in rows(scan/'Native/ArcScan.csv'):
    if row[0]=='Y' and row[1]=='1' and float(row[2])==0:
        actual=TestYield(int(row[3]),int(row[4]),accepted);yield_error=max(yield_error,abs(actual-float(row[6])))
assert yield_error<2e-11

def TestCrossings(movement):
    """На прогнозном отрезке ищет нули пересекаемых поверхностей; это геометрия прогноза, не равновесие или предел."""
    events=[]
    for point in range(1,44):
        for side in (-1,0,1):
            left=TestYield(point,side,accepted);right=TestYield(point,side,accepted+movement)
            if left>=-16e-11 or right<=0:continue
            lo,hi=0.,1.
            for _ in range(50):
                mid=(lo+hi)/2
                if TestYield(point,side,accepted+mid*movement)>0:hi=mid
                else:lo=mid
            events.append(dict(point=point,side=side,fraction=(lo+hi)/2,phi_from_MPa=left,phi_to_MPa=right))
    return sorted(events,key=lambda x:x['fraction'])

secant_events=TestCrossings(secant)
r=Reader((root/'Native/ColdBranchTangents.bin').read_bytes());assert r.raw[:16]==b'NDMCOLDTANGENT01'
meta=r.VectorTest();assert len(meta)==5 and meta[0]==reference and meta[2]==E and tuple(meta[3:])==(60,56)
results=[]
for level in (0,1/1024,1/256,1/64):
    info=r.VectorTest();field=r.VectorTest();assert info[0]==level and info[2]==info[3]==meta[1]
    np.testing.assert_array_equal(field,accepted+level*secant)
    item=dict(level=level,available=bool(info[1]))
    if info[1]:
        scales=r.VectorTest();residual=r.VectorTest();matrix=np.array([r.VectorTest() for _ in range(61)]).T
        assert len(scales)==len(residual)==61 and np.all(scales>0) and np.all(np.isfinite(matrix))
        expected=-L@source['direction']*scales[60];expected[56:]*=E
        defect=float(max(abs(matrix[:60,60]-expected)));assert defect<=1e-10*max(1,float(max(abs(expected))))
        # Нормированные строки и столбцы используются совместно: после решения поле возвращается через исходные HybridScales.
        A=matrix[:60,:60];rhs=-matrix[:60,60]/scales[60];singular=np.linalg.svd(A,compute_uv=False)
        item.update(load_column_defect=defect,condition=float(singular[0]/singular[-1]),rank_1e_minus12=int(sum(singular>singular[0]*1e-12)),
                    normed_residual=float(max(abs(residual[:60]))))
        try:
            step=np.linalg.solve(A,rhs);backward=float(max(abs(A@step-rhs))/(np.linalg.norm(A,ord=np.inf)*max(abs(step))+max(abs(rhs))))
            assert backward<1e-12
            tangent=scales[:60]*step;tangent_length=float(np.sqrt(tangent@M@tangent));assert tangent_length>0
            cosine=float(tangent@M@secant/(tangent_length*length));orientation=1. if cosine>=0 else -1.
            movement=orientation*tangent*(length/tangent_length);lambda_increment=orientation*length/tangent_length
            events=TestCrossings(movement)
            item.update(linear_backward_error=backward,cosine_with_incoming_secant=cosine,
                tangent_length_per_lambda=tangent_length,same_length_lambda_increment=lambda_increment,
                same_length_field_movement=movement.tolist(),predictor_events=events,
                predictor_is_not_accepted_equilibrium=True)
        except np.linalg.LinAlgError:item['linear_system_singular']=True
    results.append(item)
assert r.position==len(r.raw)
proof=dict(native=native,actual_components=256,only_short_readonly_entry_added=True,all_working_methods_exact858=True,
 exact19_history_and_frozen_stresses_unchanged=True,maximum_independent_yield_difference_MPa=yield_error,
 incoming_secant_length=length,incoming_secant_lambda_increment=dl,secant_predictor_events=secant_events,
 tangents=results,no_new_NDS_or_material_acceptance=True,predictors_not_equilibria=True,
 physical_capacity_certified=False,whole_goal_completed=False)
(root/'ColdBranchTangentsProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('COLD19 TANGENTS;YIELD DIFFERENCE',yield_error,'SECANT FIRST',secant_events[0] if secant_events else None)
for item in results:
    print('LEVEL',item['level'],'RANK',item.get('rank_1e_minus12'),'COND',item.get('condition'),'COS',item.get('cosine_with_incoming_secant'),
        'DL',item.get('same_length_lambda_increment'),'FIRST',(item.get('predictor_events') or [None])[0])
