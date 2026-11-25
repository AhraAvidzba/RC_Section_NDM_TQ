"""ДЛЯ ТЕСТОВ. Проверяет точную историю и независимо измеряет силы и локальность всех диагностических корней."""
from pathlib import Path
from runpy import run_path
import json,re,numpy as np
from argparse import ArgumentParser

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);group=p.add_mutually_exclusive_group();group.add_argument('--secant-guide',action='store_true');group.add_argument('--cold19',action='store_true');args=p.parse_args()
root=base/('Cold863Cold19YieldCoordinateCapture_20261125' if args.cold19 else ('Cold851LoadingSecantReparameterization_20261125' if args.secant_guide else 'Cold848NewLoadingCoordinateDiagnosis_20261125'))
parent=base/'Cold846AutomaticHGuideWalk_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin');assert len(source['fields'])==(19 if args.cold19 else 86)
if args.cold19:
    compiled=base/'Cold858BalancedInitialStressReplay_20261125';verified=base/'Cold859SpatialDeclineStressReplay_20261125'
    assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(compiled/'docs/regression/Diagnostic.xlsm').read_bytes()
    assert (root/'Input/Reverse1.bin').read_bytes()==(verified/'Input/Reverse1.bin').read_bytes()
    assert (root/'tests/modTestCoupledElasticState.bas').read_bytes()==(compiled/'tests/modTestCoupledElasticState.bas').read_bytes()
    for path in (compiled/'src').rglob('*'):
        if path.is_file():assert path.read_bytes()==(root/path.relative_to(compiled)).read_bytes(),path
    for suffix in ('.csv','.Operators.csv'):
        name='Native/EnvelopeEnergy.Accepted'+suffix;assert (root/name).read_bytes()==(verified/name).read_bytes()
    proof=json.loads((verified/'VerifiedColdStressReplayProof.json').read_bytes());assert proof['source_states']==19 and proof['source_workflow_incomplete']
    assert 'FROZEN_YIELD_DIAGNOSTIC_COMPLETED: six independent attempts; no material acceptance' in (root/'Native/RealColdReplay.txt').read_text('utf-8-sig')
if not args.cold19:
    original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
    assert len(original['fields'])==70
    for key in ('calibration','options','offset','direction'):np.testing.assert_array_equal(source[key],original[key])
    for name in ('src/Solver/CCoupledDeformationContinuation.cls',
        'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls','src/Materials/CCoupledReinforcedResponse.cls'):
        assert (root/name).read_bytes()==(parent/name).read_bytes(),name
    if not args.secant_guide:assert (root/'src/Solver/CCoupledArcContinuation.cls').read_bytes()==(parent/'src/Solver/CCoupledArcContinuation.cls').read_bytes()
    else:
        old_arc=(parent/'src/Solver/CCoupledArcContinuation.cls').read_text('utf-8-sig')
        new_arc=(root/'src/Solver/CCoupledArcContinuation.cls').read_text('utf-8-sig')
        old_pair,=re.findall(r'(?ms)^Private Function TryLoadingDeparture\(.*?^End Function',old_arc)
        new_pair,=re.findall(r'(?ms)^Private Function TryLoadingDeparturePair\(.*?^End Function',new_arc)
        assert old_pair[old_pair.index('    Dim length As Double:'):old_pair.index('    TryLoadingDeparture = True')]==new_pair[new_pair.index('    Dim length As Double:'):new_pair.index('    TryLoadingDeparturePair = True')]
    def TestMethods(path):
        """Сопоставляет рабочие методы без нормализации чисел или условий допуска."""
        rows=re.findall(r'(?ms)^((?:Private|Friend|Public) (?:Function|Sub|Property Get) (\w+)\b.*?^End (?:Sub|Function|Property))',path.read_text('utf-8-sig'))
        assert len(rows)==len({n for _,n in rows});return {n:s for s,n in rows}
    old,new=TestMethods(parent/'src/Solver/CCoupledNewtonSolver.cls'),TestMethods(root/'src/Solver/CCoupledNewtonSolver.cls')
    assert set(new)-set(old)=={'TestStartYieldCoordinateCapture','TestYieldCoordinateCapture'}
    assert {n for n in old if old[n]!=new[n]}=={'FollowYieldEventCoordinate'}
    body=new['FollowYieldEventCoordinate']
    start=body.index('Finished:\n')+len('Finished:\n');end=body.index('    mHybridActive = False',start)
    body=body[:start]+body[end:]
    body=body.replace('    Dim testStage As String: testStage = "Constraint"\n','',1)
    body=re.sub(r'^    testStage = "[A-Za-z]+"\n','',body,flags=re.M)
    assert body==old['FollowYieldEventCoordinate']
    Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
    expected_fields=list(original['fields']);expected_factors=list(original['lambdas'])
    for item in json.loads((parent/'AutomaticLoadingWalkProof.json').read_bytes())['attempts']:
        if not item['solved']:continue
        prior=Reader((parent/'Native'/f"LoadingWalk.{item['attempt']:03}.bin").read_bytes())
        packet=[prior.VectorTest() for _ in range(15)]
        expected_fields.append(prior.VectorTest());expected_factors.append(packet[0][23])
    np.testing.assert_array_equal(source['fields'],expected_fields);np.testing.assert_array_equal(source['lambdas'],expected_factors)
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
rows=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestRows']
g,L,w,v,coords,axis,b,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([b+s*.001*gradient for s in (-1,0,1)])
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1.,1.,1.,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0
stages={(int(case),int(number)):stage for case,number,stage in rows(root/'Native/YieldCoordinate.Checks.csv')}
stress={};bar={}
for row in rows(root/'Native/ElasticPrincipal.Roots.csv'):
    kind,case,point,side=row[:4];key=int(case),int(side),int(point)
    if kind=='K':stress[key]=np.array(row[4:10],float);bar[key]=float(row[10])
    else:assert kind in ('P','E')
force_scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
weight=coords[:,2];center=np.average(coords[:25,:2],axis=0,weights=weight[:25]);x,y=(coords[:,:2]-center).T

def TestMechanics(case,field,factor):
    """Из физических тензоров строит все слабые строки и шесть усилий; для отвергнутого корня фиксирует фактическую невязку."""
    net=np.zeros((3,43,6));utilization=[]
    for side in (-1,0,1):
        for point in range(1,44):
            key=case,side,point;s=stress[key];net[side+1,point-1]=s if point<=25 else bar[key]*axis[point-1]-s
            if point>25:
                axial=float((op[side+1,point-1]@field)@axis[point-1]);limit=(.025 if axial>=0 else .018) if axis[point-1,2]==1 else .02
                utilization.append(abs(axial)/limit)
    target=source['offset']+factor*source['direction'];external=L@target
    weak=g@field+np.einsum('pcd,pc->d',w,net[1])+np.einsum('pcd,pc->d',v,(net[2]-net[0])/.002)-external
    weaknorm=max(float(max(abs(weak[:56]))),float(source['calibration'][0]*max(abs(weak[56:]))))
    weighted=weak.copy();weighted[56:]*=source['calibration'][0]
    s=net[1];actual=np.array([weight@s[:,2],(weight*y)@s[:,2],(weight*x)@s[:,2],weight@(x*s[:,5]-y*s[:,4]),weight@s[:,4],weight@s[:,5]])
    bound=float(.0001*(source['options'][3]+source['options'][4]*max(abs(external[:56]))))
    delta=field-source['fields'][-1];distance=float(np.sqrt(max(0.,delta@M@delta)))
    return dict(weak_norm_MPa=weaknorm,six_force_norm_MPa=float(max(abs((actual-target)/force_scales))),
                bound_MPa=bound,distance_from_source=distance,max_steel_utilization=max(utilization),
                dominant_equation=int(np.argmax(abs(weighted)))+1,weighted_equations_MPa=weighted.tolist(),
                six_force_defect_N_Nmm=(actual-target).tolist())

r=Reader((root/'Native/YieldCoordinate.Checks.bin').read_bytes());assert r.raw[:16]==b'NDMYIELDCHECK001'
results=[];total_records=0
for case,ratio in enumerate((1,.5,.25,.125,.0625,.03125),1):
    meta=r.VectorTest();packet=[r.VectorTest() for _ in range(14)];assert meta[0]==case and meta[1]==ratio and meta[5]==meta[6]
    assert packet[0][14]==packet[0][15]==meta[5] and packet[0][18]==meta[4]
    assert packet[0][16]<=30 and packet[0][17]<=30 and packet[0][18]<=96
    assert packet[0][19]==packet[0][16]+packet[0][17]
    records=[];raw_record_vectors=[]
    for number in range(1,int(meta[7])+1):
        info,seed,scales,solution,mapped=[r.VectorTest() for _ in range(5)]
        raw_record_vectors.append((info,seed,scales,solution,mapped))
        stage=stages[case,number];entry=dict(number=number,stage=stage,info=info.tolist())
        if len(mapped)==61:
            field=(seed+scales*mapped)[:60];factor=float((seed+scales*mapped)[60]);mechanics=TestMechanics(case*10+number,field,factor)
            mechanics['distance_to_allowed_radius']=mechanics['distance_from_source']/info[7]
            entry.update(lambda_value=factor,scalar_residual_MPa=float(info[15]),
                         pivot_coordinate=float(solution[int(info[14])-1]),mechanics=mechanics)
            if info[10]:
                assert stage=='Success' and mechanics['weak_norm_MPa']<=mechanics['bound_MPa']+1e-12
                assert mechanics['six_force_norm_MPa']<=mechanics['bound_MPa']+1e-12
                assert mechanics['distance_to_allowed_radius']<=1+1e-12
            elif stage=='RadiusAndLambda':assert factor<=0 or mechanics['distance_to_allowed_radius']>1
            elif stage=='ScalarAndPivot':assert abs(info[15])>1e-11 or abs(solution[int(info[14])-1])>info[6]
            elif stage=='EquilibriumAndGauge':assert mechanics['weak_norm_MPa']>info[6]-1e-12
        records.append(entry);total_records+=1
    item=dict(case=case,ratio=ratio,arc_solved=bool(meta[2]),lambda_value=float(meta[3]),calls=int(meta[4]),
              field_half_error=float(meta[8]),coordinate_records=records)
    if meta[2]:
        field,actual,projection=[r.VectorTest() for _ in range(3)];mechanics=TestMechanics(case*10+9,field,float(meta[3]))
        assert mechanics['weak_norm_MPa']<=mechanics['bound_MPa']+1e-12 and mechanics['max_steel_utilization']<=1
        item['returned_mechanics']=mechanics
        if args.secant_guide:
            delta=field-source['fields'][-1];length=float(np.sqrt(delta@M@delta))
            half_record=records[-2];half_raw=half_record['info'];half_factor=half_raw[8]
            # Половинное поле восстанавливается из сохранённого отдельного корня с теми же битами входных данных.
            half_data=raw_record_vectors[-2];half_field=(half_data[1]+half_data[2]*half_data[4])[:60]
            error_vector=field-(2*half_field-source['fields'][-1]);field_error=float(np.sqrt(max(0.,error_vector@M@error_vector))/length)
            lambda_error=abs(2*half_factor-source['lambdas'][-1]-meta[3])
            lambda_bound=max(.005*abs(meta[3]-source['lambdas'][-1]),8*.0001*source['options'][4]*(1+abs(source['lambdas'][-1])))
            assert field_error<=.005 and lambda_error<=lambda_bound
            assert abs(field_error-meta[8])<=1e-9
            item.update(independent_field_half_error=field_error,lambda_half_error=float(lambda_error),lambda_half_bound=float(lambda_bound))
    if args.secant_guide:
        guide_reader=Reader((root/'Native'/f'LoadingGuide.{case:03}.bin').read_bytes())
        assert guide_reader.raw[:16]==b'NDMLOADGUIDE0001';present=guide_reader.VectorTest()[0]
        if present:
            guide_meta,guide_field,prediction=[guide_reader.VectorTest() for _ in range(3)]
            assert guide_meta[5]==meta[5] and guide_meta[6]==0
            accepted=source['fields'][-1];distance=float(np.sqrt((guide_field-accepted)@M@(guide_field-accepted)))
            predicted_distance=float(np.sqrt((prediction-accepted)@M@(prediction-accepted)))
            assert abs(distance-guide_meta[2])<=1e-18 and abs(predicted_distance-guide_meta[3])<=1e-18
            assert guide_meta[4]==guide_meta[3]/guide_meta[2]
            np.testing.assert_array_equal(prediction[56:],accepted[56:])
            expected=accepted.copy();expected[:56]+=guide_meta[4]*(guide_field[:56]-accepted[:56])
            np.testing.assert_array_equal(expected,prediction)
            if meta[2]:
                assert len(records)>=3 and records[-2]['stage']==records[-1]['stage']=='Success'
                assert records[-2]['info'][7]==1.5*guide_meta[3] and records[-1]['info'][7]==3*guide_meta[3]
                assert meta[4]>=guide_meta[7]+3
            item['verified_guide']=dict(meta=guide_meta.tolist(),distance=distance,predicted_distance=predicted_distance,
                                        original_radius_factors_unchanged=True,guide_not_accepted=True)
        assert guide_reader.position==len(guide_reader.raw)
    results.append(item)
assert r.position==len(r.raw) and total_records==len(stages)
proof=dict(native=native,exact_source_states=len(source['fields']),no_candidate_accepted=True,working_arc_controller_and_materials_exact846=not (args.secant_guide or args.cold19),
           all_working_sources_and_compiled_book_exact858=args.cold19,cold19_restart_not_new_cold_capacity=args.cold19,
           all_acceptance_criteria_unchanged=True,Newton_method_exact_after_removing_readonly_instrumentation=True,
           all_other_Newton_methods_exact=True,verified_guide_reparameterization=args.secant_guide,cases=results,physical_capacity_certified=False,whole_goal_completed=False)
(root/('ColdYieldCoordinateDiagnosisProof.json' if args.cold19 else ('LoadingSecantReparameterizationProof.json' if args.secant_guide else 'NewLoadingCoordinateDiagnosisProof.json'))).write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
for item in results:
    print('RATIO',item['ratio'],'SOLVED',item['arc_solved'],'ERROR',item['field_half_error'])
    for entry in item['coordinate_records']:
        print(entry['number'],entry['stage'],'stop',entry['info'][11],'solver residual',entry['info'][12],
              'phi',entry.get('scalar_residual_MPa'),'pivot',entry.get('pivot_coordinate'),'mechanics',entry.get('mechanics'))
