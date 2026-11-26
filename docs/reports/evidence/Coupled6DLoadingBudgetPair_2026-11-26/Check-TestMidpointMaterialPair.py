"""ДЛЯ ТЕСТОВ. Проверяет последовательную материальную пару974, нативный средний прогноз и действительный общий бюджет30."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,default=base/'Cold974MidpointMaterialPair_20261126')
root=parser.parse_args().root;parent=base/'Cold973OutgoingMaterialPair_20261126'
if root.name=='Cold976PublishedMaterialPair_20261126':parent=base/'Cold975ResolvedOutgoingPair_20261126'
if root.name in ('Cold977MaterialDirectionAudit_20261126','Cold978LoadingMaterialDirection_20261126'):
    parent=base/'Cold976PublishedMaterialPair_20261126'
loading_budget=root.name=='Cold980LoadingBudgetPair_20261126'
if loading_budget:parent=base/'Cold979LoadingAlgorithmComparison_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==257 and audit['imported_this_pass']==['modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==256
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
methods=run_path(str(here/'Stage-TestWorkingDepartureTrace.py'))['TestMethods']
before=methods((parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig'))
after=methods((root/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig'))
added='TestRunPublishedMaterialPair' if root.name=='Cold976PublishedMaterialPair_20261126' else 'TestRunMaterialCoordinatePair'
if root.name in ('Cold977MaterialDirectionAudit_20261126','Cold978LoadingMaterialDirection_20261126'):
    added='TestRunMaterialDirectionAudit'
if loading_budget:
    assert set(after)==set(before)
    changed=('Function','TestRunPublishedMaterialPair')
    assert before[changed].group().replace('If caseIndex = 1 Then budget = budget \\ 2','If caseIndex = 1 Then budget = (2 * budget) \\ 3',1)==after[changed].group()
    assert all(before[key].group()==after[key].group() for key in before if key!=changed)
else:
    assert set(after)-set(before)=={('Function',added)}
    assert all(before[key].group()==after[key].group() for key in before)
if root.name in ('Cold974MidpointMaterialPair_20261126','Cold976PublishedMaterialPair_20261126','Cold977MaterialDirectionAudit_20261126'):
    assert (root/'Input/RecordedYieldSeed.bin').read_bytes()==(parent/'Input/RecordedYieldSeed.bin').read_bytes()
assert (root/'Input/Reverse1.bin').read_bytes()==(parent/'Input/Reverse1.bin').read_bytes()
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu))
D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
r=R((root/'SourceAcceptedSurfaceStates.bin').read_bytes());states={}
for _ in range(129):
    point,side=r.VectorTest();states[int(point),int(side)]=r.VectorTest()
assert r.position==len(r.raw)
helper=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig')
begin=helper.index('def TestYield(');end=helper.index('summary=[];',begin)
body=helper[begin:end];old='return dict(residual=residual,concrete=concrete';assert old in body
body=body.replace(old,'return dict(six_force_residual_vector=(actual-target)/scales,residual=residual,concrete=concrete',1)
exec(compile(body,str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())

def TestLength(delta):
    """Восстанавливает общую тензорную меру по трём станциям, не включая служебные множители."""
    strain=np.einsum('spcd,d->spc',op,delta)
    return float(np.sqrt(np.einsum('spc,spc,p,c->',strain,strain,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]))))

r=R((root/'Native/MaterialMidpointPair.bin').read_bytes());assert r.raw[:16]==b'NDMMATPAIR000001'
meta,context,accepted,projection=[r.VectorTest() for _ in range(4)]
np.testing.assert_array_equal(accepted,source['fields'][-1]);reference=float(context[0]);revision=context[1]
assert reference==source['lambdas'][-1] and meta[5:8].tolist()==[source['options'][3],source['options'][4],source['options'][2]]
if root.name=='Cold975ResolvedOutgoingPair_20261126':
    old_root=base/'Cold974MidpointMaterialPair_20261126'
    assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(old_root/'docs/regression/Diagnostic.xlsm').read_bytes()
    old_reader=R((old_root/'Input/RecordedYieldSeed.bin').read_bytes());old_reader.position=0
    old_meta,old_seed,old_accepted,old_residual,old_projection=[old_reader.VectorTest() for _ in range(5)]
    new_reader=R((root/'Input/RecordedYieldSeed.bin').read_bytes());new_reader.position=0
    new_meta,new_seed,new_accepted,new_residual,new_projection=[new_reader.VectorTest() for _ in range(5)]
    for a,b in ((old_seed,new_seed),(old_accepted,new_accepted),(old_residual,new_residual),(old_projection,new_projection)):
        np.testing.assert_array_equal(a,b)
    for i in (0,1,3,4,5,6,7,8):assert old_meta[i]==new_meta[i]
    value=TestYield(int(meta[0]),int(meta[1]),accepted)
    assert meta[2]==value+16*(old_meta[2]-value) and meta[9]==.5*(value+meta[2])
cases=[];fields=[];lambdas=[];total=0;tag=r.VectorTest()
while tag[0]>0:
    index=int(tag[0]);assert index==len(cases)+1 and len(tag)==4 and tag[2]==total
    budget=int(tag[1]);assert budget==((20 if loading_budget else 15) if index==1 else 30-total)
    seed,initial_values,info=[r.VectorTest() for _ in range(3)];assert len(info)==15 and info[0]==index and info[2]==1 and info[10]<=budget
    if index==2:
        np.testing.assert_array_equal(seed,.5*(accepted+fields[0]))
        assert tag[3]==.5*(reference+lambdas[0])
    total+=int(info[10]);assert total<=30
    item=dict(case=index,budget=budget,steps=int(info[10]),total_steps=total,seed_lambda=float(tag[3]),
        target_MPa=float(info[1]),native_solved=bool(info[3]),material_ready=bool(info[5]),lambda_value=float(info[6]),
        native_coordinate_residual_MPa=float(info[7]),native_norm_MPa=float(info[9]),calls=int(info[11]),
        jacobians=int(info[12]),stop_code=int(info[14]),independently_equilibrated=False)
    assert info[1]==(meta[2] if index==1 else meta[9])
    if info[5]:
        field,values,force_residual=[r.VectorTest() for _ in range(3)];stress=np.zeros((3,43,6));seen=set()
        for _ in range(129):
            point,side,code=r.VectorTest();sigma=r.VectorTest();key=int(point),int(side)
            assert code==0 and key not in seen;seen.add(key);stress[int(side)+1,int(point)-1]=sigma
        assembled=TestAssemble(field,info[6],stress)
        assert max(abs(values[:56]-assembled['residual'][:56]))<2e-11
        assert max(abs(values[56:60]-E*assembled['residual'][56:]))<2e-11
        assert max(abs(force_residual-assembled['six_force_residual_vector']))<2e-11
        bound=float(meta[5]+meta[6]*max(abs((L@(source['offset']+info[6]*source['direction']))[:56])))
        norm=float(max(assembled['physical_norm'],assembled['six_force_norm'],E*max(abs(assembled['residual'][56:]))))
        assert abs(norm-info[9])<2e-11
        coordinate=TestYield(int(meta[0]),int(meta[1]),field)-info[1];difference=abs(float(coordinate-info[7]));assert difference<2e-10
        strain=np.einsum('spcd,d->spc',op,field);bars=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
        utilization=float(np.max(abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)))
        assert abs(values[-1]-E*(projection@field-meta[4]))<2e-11
        valid=bool(norm<=bound+1e-12 and abs(info[7])<=2e-12 and abs(coordinate)<=2e-12+difference and utilization<=1 and info[6]>0)
        item.update(independently_equilibrated=valid,residual_MPa=norm,bound_MPa=bound,
            physical_residual_MPa=assembled['physical_norm'],six_force_residual_MPa=assembled['six_force_norm'],
            gauge_residual_MPa=float(E*max(abs(assembled['residual'][56:]))),
            worst_physical_equation_1based=int(np.argmax(abs(assembled['residual'][:56])))+1,
            independent_coordinate_residual_MPa=float(coordinate),coordinate_reconstruction_difference_MPa=difference,
            steel_utilization=utilization)
        fields.append(field);lambdas.append(float(info[6]))
    cases.append(item);tag=r.VectorTest()
assert tag.tolist()==[0.,float(total),float(revision)] and r.position==len(r.raw)
pair=None
if len(cases)==2 and all(item['independently_equilibrated'] for item in cases):
    full,half=fields;length=TestLength(full-accepted);error=TestLength(full-(2*half-accepted))/length
    alpha=json.loads((root/'SourceAutomatic962ArcProof.json').read_bytes())['info'][9]
    lambda_error=abs(2*lambdas[1]-reference-lambdas[0]);radius=float(np.hypot(length,alpha*(lambdas[0]-reference)))
    allowance=(.005*radius)**2-(length*error)**2;noise=8*source['options'][4]*(1+abs(reference))
    bound=max(noise,float(np.sqrt(allowance)/alpha) if allowance>0 else noise)
    resolution=16*cases[0]['bound_MPa']/E
    pair=dict(field_error=error,lambda_error=lambda_error,lambda_bound=float(bound),
        joint_error=float(np.hypot(length*error,alpha*lambda_error)/radius),field_length=length,
        material_target_half_of_original_increment=bool(abs(meta[9]-.5*(TestYield(int(meta[0]),int(meta[1]),accepted)+meta[2]))<1e-16),
        passes_existing_independent_half_limits=bool(error<=.005 and lambda_error<=bound),
        existing_spatial_field_resolution=resolution,resolved_spatial_field_advance=bool(length>resolution),
        passes_existing_shape_and_resolution_limits=bool(error<=.005 and lambda_error<=bound and length>resolution),
        full_lambda=lambdas[0],half_lambda=lambdas[1],full_plus_half_steps=total)
proof=dict(native=native,actual_components=257,all256_numerical_components_preserved=True,only_new_test_entry_added=not loading_budget,
    only_test_phase_budget15_to20=loading_budget,
    saved_book_bitexact974_and_only_numerical_target_increment16x=bool(root.name=='Cold975ResolvedOutgoingPair_20261126'),
    original25_history_and_revision_unchanged_by_native_guards=True,original_force_path_and_materials_and_tolerances_preserved=True,
    first_initial_residual_bitexact_actual968_guard=True,half_seed_and_lambda_midpoint_verified_native=True,
    actual_shared_step_budget=30,total_steps=total,cases=cases,pair=pair,
    native_lambda_uncertainty_not_recorded=True,no_material_history_acceptance=True,
    automatic_station_selection_cost_not_included_here=True,not_production_recovery_acceptance=True,
    physical_capacity_certified=False,cold_gain_claimed=False,whole_goal_completed=False)
(root/'MidpointMaterialPairProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('MIDPOINT MATERIAL PAIR',cases,'PAIR',pair,'TOTAL STEPS',total)
