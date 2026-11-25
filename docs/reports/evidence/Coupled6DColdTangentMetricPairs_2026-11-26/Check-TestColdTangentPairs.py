"""ДЛЯ ТЕСТОВ. Независимо проверяет касательные пары, материальное продвижение и исходные критерии при неизменной холодной истории19."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
from argparse import ArgumentParser

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,default=base/'Cold865Cold19TangentPairs_20261125');args=p.parse_args()
root=args.root;parent=base/'Cold858BalancedInitialStressReplay_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and audit['imported_this_pass']==['modTestCoupledElasticState'] and len(audit['other_actual_components_exact_parent_except_identifier_case'])==255
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes(),path
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];source=read(root/'Input/Reverse1.bin');assert len(source['fields'])==19
for suffix in ('.csv','.Operators.csv'):
    assert (root/('Native/ColdTangent.Restored.Accepted'+suffix)).read_bytes()==(root/('Native/EnvelopeEnergy.Accepted'+suffix)).read_bytes()
validate=run_path(str(here/'Validate-TestHardPrepared.py'));input_mechanics,input_fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(input_fields,source['fields'],rtol=1e-14,atol=1e-300)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0
accepted=source['fields'][-1];secant=accepted-source['fields'][-2];reference=float(source['lambdas'][-1]);E=source['calibration'][0]

def TestLength(delta):
    """Использует ту же тензорную метрику трёх станций без служебных множителей."""
    return float(np.sqrt(max(0.,delta@M@delta)))

def TestBound(factor):
    """Восстанавливает строгий исходный допуск на фактической нагрузке, без изменения единиц и знаков."""
    return float(.0001*(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+factor*source['direction']))[:56]))))

def TestEnergy(phase,count):
    """Одним проходом суммирует энергетические координаты реальных бетонных точек последнего состояния, не подменяя ими повреждение."""
    total=0.;seen=set();gt,gc,ell=source['calibration'][6:9]
    for row in validate['TestRows'](root/'Native'/(phase+'.csv')):
        if row[0]=='S' and int(row[1])==count and int(row[3])<=25:
            key=(int(row[2]),int(row[3]));assert key not in seen;seen.add(key)
            total+=coords[key[1]-1,2]*(gt*float(row[10])+gc*float(row[11]))/ell
    assert len(seen)==75
    return total/(3*sum(coords[:25,2]))

baselineH=TestEnergy('EnvelopeEnergy.Accepted',19)
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=Reader((root/'Native/ColdTangentPairs.bin').read_bytes());assert r.raw[:16]==b'NDMCOLDPAIR00001'
header=r.VectorTest();assert header[0]==3 and header[1]==reference
slopes=[TestLength(source['fields'][i+1]-source['fields'][i])/abs(source['lambdas'][i+1]-source['lambdas'][i]) for i in range(15,18)]
alpha=float(np.median(slopes));reference_metric=None
if (root/'Native/ColdReferenceMetric.bin').is_file():
    ref=Reader((root/'Native/ColdReferenceMetric.bin').read_bytes());assert ref.raw[:16]==b'NDMCOLDREFMET001'
    info=ref.VectorTest();direction=ref.VectorTest();linear_field=ref.VectorTest();recovered=ref.VectorTest();assert ref.position==len(ref.raw)
    assert abs(info[1]-alpha)<=1e-15*alpha and info[2]==E
    np.testing.assert_array_equal(direction,source['direction'])
    independent_alpha=TestLength(linear_field);assert abs(independent_alpha-info[0])<=1e-15*independent_alpha
    # Упругая матрица заново строится из опубликованных операторов; сталь даёт направленный вклад и вычитает бетон замещения.
    nu=source['calibration'][1];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu));D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
    tensors=np.tile(D,(43,1,1));tensors[25:]=200000*np.einsum('pc,pd->pcd',axis[25:],axis[25:])-D
    elastic=g+np.einsum('pcr,pcd,pds->rs',w,tensors,B)+np.einsum('pcr,pcd,pds->rs',v,tensors,gradient)
    defect=elastic@linear_field-L@direction
    norm=max(max(abs(defect[:56])),E*max(abs(defect[56:])))
    assert norm<1e-8
    assert max(abs(recovered-direction))<1e-5
    reference_metric=dict(alpha=float(info[0]),previous_median_alpha=alpha,independent_elastic_residual_MPa=float(norm),
                          linear_field=linear_field.tolist(),prepared_elastic_direction_recovered=True)
    alpha=float(info[0])
assert abs(alpha-header[3])<=1e-15*alpha
inputs=Reader((root/'Input/ColdTangentPredictors.bin').read_bytes());inputs.position=0;assert inputs.VectorTest()[0]==3
cases=[(inputs.VectorTest(),inputs.VectorTest()) for _ in range(3)];assert inputs.position==len(inputs.raw)
results=[]
for case,(request,movement) in enumerate(cases,1):
    meta=r.VectorTest();vectors=[r.VectorTest() for _ in range(12)]
    projection,stored,half_seed,full_seed,half,full,actual_half,actual_full,half_effort,full_effort,restored,restored_actual=vectors
    assert len(meta)==23 and meta[0]==case and meta[1]==request[0] and meta[2]==request[1] and meta[4]==request[2]
    assert meta[3]==reference and meta[19]==meta[20]==header[2]
    np.testing.assert_array_equal(stored,accepted)
    ratio=float(meta[2]);dl=float(meta[4]);spacing=TestLength(movement);arc=float(np.sqrt(spacing**2+alpha**2*dl**2));radius=ratio*arc
    assert abs(meta[6]-spacing)<=1e-18 and abs(meta[7]-arc)<=1e-18 and abs(meta[8]-radius)<=1e-18
    expected_projection=M@movement/arc;q=alpha**2*dl/arc
    np.testing.assert_allclose(projection,expected_projection,rtol=1e-10,atol=1e-16)
    assert abs(meta[9]-q)<=1e-18
    expected_seed=accepted.copy();expected_seed[:56]+=.5*ratio*movement[:56]
    np.testing.assert_array_equal(half_seed,expected_seed)
    sh,sf=bool(meta[10]),bool(meta[11]);lh,lf=map(float,meta[12:14]);half_steps,full_steps=map(int,meta[14:16])
    assert half_steps+full_steps<=source['options'][2] and half_steps==half_effort[0] and full_steps==full_effort[0]
    item=dict(case=case,level=float(meta[1]),ratio=ratio,half_solved=sh,full_solved=sf,half_lambda=lh,full_lambda=lf,
              pair_iterations=half_steps+full_steps,source_revision_unchanged=True,roots={})
    for role,ready,field,factor,reported_util,actual in (('Half',sh,half,lh,meta[21],actual_half),('Full',sf,full,lf,meta[22],actual_full)):
        if not ready:
            assert len(field)==1;continue
        phase='ColdTangent.'+role+str(case).zfill(3)+'.Sequence';mechanics,fields=validate['TestValidate'](root,phase)
        assert len(fields)==20
        np.testing.assert_allclose(fields[:19],source['fields'],rtol=1e-14,atol=1e-300)
        np.testing.assert_allclose(fields[-1],field,rtol=1e-14,atol=1e-300)
        weak=mechanics['states'][-1];bound=TestBound(factor)
        assert weak['residual_MPa']<=bound+1e-12 and weak['force_residual_MPa']<=bound+1e-12
        target=source['offset']+factor*source['direction'];scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
        assert max(abs((actual-target)/scales))<=bound
        strain=np.einsum('spcd,d->spc',op,field);steel=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
        limit=np.where(long[None,:],np.where(steel>=0,.025,.018),.02);util=float(np.max(abs(steel)/limit))
        assert util<=1 and abs(util-reported_util)<=1e-12
        h=TestEnergy(phase,20);length=TestLength(field-accepted);distance=float(np.sqrt(length**2+alpha**2*(factor-reference)**2))
        fraction=.5 if role=='Half' else 1.;control=projection@(field-accepted)+q*(factor-reference)
        assert abs(control-fraction*radius)*E<=bound+1e-12
        local=factor>0 and factor<=source['options'][0] and distance<=3*fraction*radius and control>0
        item['roots'][role]=dict(weak=weak,strict_bound_MPa=bound,steel_utilization=util,H_MPa=h,H_increment_MPa=h-baselineH,
            length=length,combined_distance=distance,original_locality=bool(local),lambda_decreased=factor<reference)
    if sh and sf:
        length=TestLength(full-accepted);field_error=TestLength(full-(2*half-accepted))/length
        lambda_error=abs(2*lh-reference-lf);lambda_bound=max(.005*abs(lf-reference),8*.0001*source['options'][4]*(1+abs(reference)))
        combined_length=float(np.sqrt(length**2+alpha**2*(lf-reference)**2))
        combined_error=float(np.sqrt((field_error*length)**2+alpha**2*lambda_error**2)/combined_length)
        cosine=float((full-accepted)@M@secant/(length*TestLength(secant)))
        restored_ok=bool(meta[16]);assert restored_ok and meta[17]==0 and meta[18]==lf
        np.testing.assert_array_equal(restored,full);np.testing.assert_array_equal(restored_actual,actual_full)
        resolved=.5*radius*E>8*max(TestBound(reference),TestBound(lh),TestBound(lf))
        growth=all(x['H_increment_MPa']>0 for x in item['roots'].values())
        qualifies=field_error<=.005 and lambda_error<=lambda_bound and cosine>0 and resolved and growth and all(x['original_locality'] for x in item['roots'].values())
        qualifies_joint=field_error<=.005 and combined_error<=.005 and cosine>0 and resolved and growth and all(x['original_locality'] for x in item['roots'].values())
        item.update(field_half_error=field_error,lambda_half_error=lambda_error,original_lambda_half_bound=lambda_bound,
            combined_arclength_half_error_diagnostic=combined_error,combined_arclength_threshold_diagnostic=.005,
            direction_cosine=cosine,material_H_advances=bool(growth),original_resolution=bool(resolved),
            restoration_zero_iterations_same_lambda_bit_identical=True,qualifies_original_arc_criteria=bool(qualifies),
            qualifies_joint_metric_diagnostic=bool(qualifies_joint))
    results.append(item)
assert r.position==len(r.raw)
proof=dict(native=native,actual_components=256,all_working_numerical_and_material_methods_exact858=True,
 true_source19_history_and_final_restored_csv_bit_identical=True,no_future_field_or_artificial_history=True,
 baseline_H_MPa=baselineH,cases=results,original_force_and_steel_criteria_unchanged=True,reference_metric=reference_metric,
 combined_arclength_metric_only_diagnostic_not_admission=True,no_branch_or_capacity_acceptance=True,
 physical_capacity_certified=False,whole_goal_completed=False)
(root/'ColdTangentPairsProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
for item in results:
    print('CASE',item['case'],'ROOTS',item['half_solved'],item['full_solved'],'ITERS',item['pair_iterations'],
          'FIELDERR',item.get('field_half_error'),'LAMBDAERR/BOUND',item.get('lambda_half_error'),item.get('original_lambda_half_bound'),
          'COMBINEDERR',item.get('combined_arclength_half_error_diagnostic'),'H ADVANCES',item.get('material_H_advances'),
          'ORIGINAL ADMISSION',item.get('qualifies_original_arc_criteria'),'JOINT DIAGNOSTIC',item.get('qualifies_joint_metric_diagnostic'))
