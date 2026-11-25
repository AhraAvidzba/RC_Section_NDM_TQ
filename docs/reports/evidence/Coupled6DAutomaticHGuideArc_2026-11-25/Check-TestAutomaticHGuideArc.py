"""ДЛЯ ТЕСТОВ. Независимо проверяет автоматический H-переход, остатки вызовов и сохранность исходной математики материальной пары."""
from pathlib import Path
from runpy import run_path
import json,re,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold844AutomaticHGuideArc_20261125';parent=base/'Cold842HReentryWalk_20261125';source_root=base/'Cold841HGuideMaterialEvent_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls','src/Materials/CCoupledConcreteReturn.cls',
    'src/Solver/CCoupledMaterialConstraint.cls','src/Solver/CCoupledEnvelopeEquations.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
assert (root/'Input/Reverse1.bin').read_bytes()==(source_root/'Input/Reverse1.bin').read_bytes()
assert not (root/'Input/HMaterialGuide.bin').exists()
assert 'Input/HMaterialGuide.bin' not in json.loads((root/'Preparation.json').read_bytes())['inputs']
def TestMethods(code):
    """Сопоставляет тела методов без нормализации арифметики и условий."""
    rows=re.findall(r'(?ms)^((?:Private|Friend|Public) (?:Function|Sub|Property Get) (\w+)\b.*?^End (?:Sub|Function|Property))',code)
    assert len(rows)==len({name for _,name in rows});return {name:body for body,name in rows}
name='src/Solver/CCoupledArcContinuation.cls';a,b=TestMethods((parent/name).read_text('utf-8-sig')),TestMethods((root/name).read_text('utf-8-sig'))
assert set(b)-set(a)=={'IsEnvelopeUnloading','TryEnvelopeGuidedEvent','EnvelopeGuideVerification'}
assert {n for n in a if a[n]!=b[n]}=={'TryStep','TryEvent'}
new=b['TryEvent'].replace('ByRef selectedKind As Long, _\n        Optional ByVal spentIterations As Long = 0) As Boolean',
    'ByRef selectedKind As Long) As Boolean')
new=new.replace('    If mMaximumIterations - spentIterations < 2 Then Exit Function\n','',1)
new=new.replace('(mMaximumIterations - spentIterations) \\ 2','mMaximumIterations \\ 2',1)
new=new.replace('    Dim eventSteps As Long: eventSteps = mSolver.IterationCount\n    phaseSteps = spentIterations + eventSteps: mPrimary = mPrimary + eventSteps',
    '    phaseSteps = mSolver.IterationCount: mPrimary = mPrimary + phaseSteps',1)
assert new==a['TryEvent'],'Force/half/lambda/event mathematics must remain unchanged'
old_pair=a['TryStep'].split('    If sf And sh Then\n',1)[1].split('    Dim candidate()',1)[0]
new_pair=b['TryStep'].split('    If sf And sh Then\n',1)[1].split('    Dim unloading As Boolean',1)[0]
assert old_pair==new_pair
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=Reader((root/'Native/AutomaticHGuide.bin').read_bytes())
assert r.raw[:16]==b'NDMAUTOH00000001'
records=[];mechanical_cases={}
for case in range(1,4):
    meta=r.VectorTest();packet=[r.VectorTest() for _ in range(15)]
    field,actual,projection,effort,present=[r.VectorTest() for _ in range(5)]
    guide=[r.VectorTest() for _ in range(3)] if present[0] else None
    assert meta[0]==case and meta[1]==(5,6,96)[case-1] and meta[4]==meta[5]==packet[0][14]==packet[0][15]
    info=packet[0];assert info[16]<=30 and info[17]<=30 and info[19]==info[16]+info[17]
    assert info[18]<=meta[1] and bool(info[3])==bool(meta[2])
    if case==1:
        assert not meta[2] and info[18]==2 and guide is None
        assert info[0]==1 and info[1]==0  # Отсутствующий половинный корень не читается и не заменяется фиктивным результатом.
        mechanical_cases[11]=dict(field=packet[6],lambda_value=float(info[4]),accuracy=.0001)
    else:
        assert meta[2]==1 and info[18]==6 and info[28]==2 and info[26]==1
        assert guide is not None and guide[0][0]==1 and guide[0][1]==1
        assert effort[0]==0 and info[23]==info[27]==meta[3]
        np.testing.assert_array_equal(field,packet[10]);assert len(packet[9])==7
        mechanical_cases[10*case+1]=dict(field=field,lambda_value=float(info[23]),accuracy=.0001,actual=actual)
        mechanical_cases[10*case+2]=dict(field=packet[11],lambda_value=float(info[24]),accuracy=.0001)
        mechanical_cases[10*case+3]=dict(field=guide[2],lambda_value=float(guide[0][6]),accuracy=.0001)
    records.append(dict(meta=meta,packet=packet,field=field,actual=actual,projection=projection,effort=effort,guide=guide))
assert r.position==len(r.raw)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
assert len(source['fields'])==70
independent,M=run_path(str(here/'TestUnacceptedMechanics.py'))['TestUnacceptedRoots'](root,source,mechanical_cases)
validate=run_path(str(here/'Validate-TestHardPrepared.py'));source_mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
g,L,w,v,coords,axis,basis,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
rows=validate['TestRows'](root/'Native/EnvelopeEnergy.Accepted.csv');stored=np.zeros((3,43,2))
for row in rows:
    if row[0]=='S' and int(row[1])==70:stored[int(row[2])+1,int(row[3])-1]=np.array(row[10:12],float)
E,nu=source['calibration'][:2];gt,gc,localization=source['calibration'][6:9]
source_h=float(np.einsum('spc,p,c->',stored[:,:25],coords[:25,2],np.array([gt/localization,gc/localization]))/(3*sum(coords[:25,2])))
reference=source['fields'][-1];reference_lambda=float(source['lambdas'][-1])
assert abs(independent[11]['independent_point_energy_average_MPa']-source_h)<=2e-18
assert records[0]['packet'][0][4] < reference_lambda-8*source['options'][4]*(1+abs(reference_lambda))
def TestLength(delta):
    """Независимая тензорная метрика трёх станций без координат жёстких связей."""
    return float(np.sqrt(max(0.,delta@M@delta)))
off=E*nu/((1+nu)*(1-2*nu));G=E/(2*(1+nu));D=np.diag([off+2*G]*3+[G]*3);D[:3,:3]+=off*(np.ones((3,3))-np.eye(3))
results=[]
for case in (2,3):
    item=records[case-1];packet=item['packet'];info=packet[0];event=packet[9];state=packet[14];guide_meta,seed,guide=item['guide']
    assert guide_meta[2]<=15 and guide_meta[14]==guide_meta[15]==info[14]
    assert info[17]>=guide_meta[2] and info[18]==2+guide_meta[1]+2+1
    assert abs(guide_meta[3]-source_h)<=2e-18
    assert independent[10*case+3]['independent_point_energy_average_MPa']>source_h
    assert abs(independent[10*case+3]['independent_point_energy_average_MPa']-guide_meta[8])<=1e-16
    assert abs(guide_meta[8]-guide_meta[5])*(E/source['calibration'][3])<=independent[10*case+3]['bound_MPa']+1e-12
    face=item['field'];half=packet[11];length=TestLength(face-reference)
    field_error=TestLength(face-(2*half-reference))/length
    lambda_error=float(abs(2*info[24]-reference_lambda-info[23]))
    lambda_bound=max(.005*abs(info[23]-reference_lambda),8*.0001*source['options'][4]*(1+abs(reference_lambda)))
    assert field_error<=.005 and lambda_error<=lambda_bound and abs(field_error-info[25])<=1e-9
    assert length<=guide_meta[16] and guide_meta[16]==3*info[11]
    assert .5*length*E>8*max(independent[10*case+1]['bound_MPa'],independent[10*case+2]['bound_MPa'])
    assert source_h<independent[10*case+2]['independent_point_energy_average_MPa']<independent[10*case+1]['independent_point_energy_average_MPa']
    point,side=map(int,event[1:3]);operator=basis[point-1]+side*.001*gradient[point-1]
    def TestYield(field):
        """Восстанавливает поверхность выбранной станции из принятой пластической истории и упругого тензора."""
        alpha,gamma=state[:2];cc,ct=state[8:];xx,yy,zz,xy,xz,yz=D@(operator@field-state[2:8])
        tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);trace=np.trace(tensor);high=np.linalg.eigvalsh(tensor)[-1]
        q=np.sqrt(1.5*np.sum((tensor-np.eye(3)*trace/3)**2));term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
        return float((alpha*trace+q+term)/(1-alpha)-cc)
    phi0,phi_guide,phi_face,phi_half=map(TestYield,(reference,guide,face,half))
    assert phi0<0 and phi_guide>0 and phi_half<0 and abs(phi_face)<=6e-12
    assert abs(phi0-event[3])<=2e-10 and abs(phi_guide-event[4])<=2e-10
    assert abs(TestYield(reference+event[5]*(guide-reference)))<=6e-12
    results.append(dict(allowance=int(item['meta'][1]),calls=int(info[18]),pair_primary=int(info[16]),
        recovery_primary=int(info[17]),total_primary=int(info[19]),guide_primary=int(guide_meta[2]),
        event_point=point,event_side=side,event_lambda=float(info[23]),field_half_error=field_error,
        lambda_half_error=lambda_error,lambda_half_bound=float(lambda_bound),source_H_MPa=source_h,
        half_H_MPa=independent[10*case+2]['independent_point_energy_average_MPa'],
        event_H_MPa=independent[10*case+1]['independent_point_energy_average_MPa'],
        mechanics={str(k):independent[k] for k in (10*case+1,10*case+2,10*case+3)},
        event_phi_MPa=phi_face,predeclared_radius=float(guide_meta[16]),event_distance=length))
np.testing.assert_array_equal(records[1]['field'],records[2]['field'])
np.testing.assert_array_equal(records[1]['guide'][2],records[2]['guide'][2])
proof=dict(native=native,source70_equilibria_verified=source_mechanics,original_material_and_Newton_math_unchanged=True,
    original_pair_and_event_force_half_lambda_math_unchanged=True,all_original_other_arc_methods_unchanged=True,
    no_precomputed_future_field_available_at_expected_input_path=True,
    allowance5_unloading_root_not_admitted_with_two_calls=True,allowance6_and96_return_bit_identical_event=True,
    guide_event_half_and_restore_calls_all_counted=True,recovery_phase_includes_guide_and_is_at_most30=True,
    no_candidate_history_accepted=True,successful_cases=results,
    physical_capacity_certified=False,whole_goal_completed=False,actual_cold_integration_not_yet_run=True)
(root/'AutomaticHGuideArcProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('AUTOMATIC H GUIDE VERIFIED:5-CALL CONTROL REJECTED;6/96 BIT IDENTICAL;ALL FORCES/HALF/STEEL/GUIDE BUDGETS VERIFIED')
for item in results:print(item)
