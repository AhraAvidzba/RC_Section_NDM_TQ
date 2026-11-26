"""ДЛЯ ТЕСТОВ. Независимо пересобирает984 до/после commit и каждую пару на фактически обновлённых бетонных состояниях."""
from pathlib import Path
from runpy import run_path
import json,sys,struct,numpy as np
from argparse import ArgumentParser

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,default=base/'Cold984AcceptedLoadingHistory_20261126')
root=parser.parse_args().root;parent=base/'Cold980LoadingBudgetPair_20261126';growing=root.name=='Cold985GrowingLoadingHistory_20261126'
sys.argv=[str(here/'Check-TestMaterialDirectionAudit.py'),'--root',str(root)]
run_path(str(here/'Check-TestMaterialDirectionAudit.py'),run_name='accepted_loading_first_pair_direction_test')
sys.argv=[str(here/'Check-TestMidpointMaterialPair.py'),'--root',str(root)]
env=run_path(str(here/'Check-TestMidpointMaterialPair.py'),run_name='accepted_loading_pair_operators_test')
R=env['R'];source=env['source'];E=env['E'];L=env['L'];coords=env['coords'];op=env['op'];axis=env['axis']
assert (root/'Native/MaterialMidpointPair.bin').read_bytes()==(parent/'Native/MaterialMidpointPair.bin').read_bytes()
assert (root/'Input/RecordedYieldSeed.bin').read_bytes()==(parent/'Input/RecordedYieldSeed.bin').read_bytes()
trajectory=list(source['fields']);lambdas=list(source['lambdas'])

def TestState(path):
    """Читает точную принятую историю всех129 точек, включая замещаемый бетон в местах арматуры."""
    r=R(path.read_bytes());assert r.raw[:16]==b'NDMHISTSTATE0001'
    revision=r.VectorTest();field=r.VectorTest();records={}
    for _ in range(129):
        point,side=r.VectorTest();key=int(point),int(side);assert key not in records
        state,damage,stress=[r.VectorTest() for _ in range(3)]
        assert len(state)==10 and len(damage)==3 and len(stress)==6
        records[key]=dict(state=state,damage=damage,stress=stress)
    assert r.position==len(r.raw)
    return int(revision[0]),field,records

def TestStress(r):
    """Читает полный пространственный тензор каждой точки трёх станций без потери точности Double."""
    stress=np.zeros((3,43,6));seen=set()
    for _ in range(129):
        point,side,code=r.VectorTest();sigma=r.VectorTest();key=int(point),int(side)
        assert code==0 and key not in seen;seen.add(key);stress[int(side)+1,int(point)-1]=sigma
    return stress

def TestSystem(field,factor,stress):
    """Собирает все56 физических условий, четыре связи и шесть результирующих при исходных единицах и допусках."""
    result=env['TestAssemble'](field,factor,stress)
    bound=float(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+factor*source['direction']))[:56])))
    norm=float(max(result['physical_norm'],result['six_force_norm'],E*max(abs(result['residual'][56:]))))
    strain=np.einsum('spcd,d->spc',op,field)
    bars=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
    utilization=float(np.max(abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)))
    scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
    result['actual_forces']=source['offset']+factor*source['direction']+result['six_force_residual_vector']*scales
    result['force_scales']=scales
    return dict(norm=norm,bound=bound,steel_utilization=utilization,verified=norm<=bound+1e-12 and utilization<=1),result

def TestArcLength(delta):
    """Восстанавливает штатную меру ArcPhysicalSquare по постоянным и нечётным коэффициентам деформаций."""
    strains=np.einsum('spcd,d->spc',op,delta)
    a=strains[1];b=.5*(strains[2]-strains[0]);weights=np.array([1,1,1,.5,.5,.5])
    square=np.einsum('pc,pc,p,c->',a,a,coords[:,2],weights)
    square+=(2/3)*np.einsum('pc,pc,p,c->',b,b,coords[:,2],weights)
    return float(np.sqrt(square/sum(coords[:,2])))

def TestH(records):
    """Воспроизводит H только по25 бетонным весам и сохранённым долям огибающих; это не энергия разрушения всего сечения."""
    total=0.;gt,gc,length=source['calibration'][6:9]
    for point in range(1,26):
        for side in (-1,0,1):
            tension,compression=records[point,side]['damage'][:2]
            total+=coords[point-1,2]*(gt/length*tension+gc/length*compression)
    return float(total/(3*sum(coords[:25,2])))

def TestPublication(folder,pair,fields,meta,revision):
    """Проверяет каждую опубликованную пару по фактическим сопряжённым матрицам и независимому решению для оценки ошибки lambda."""
    r=R((folder/'Native/MaterialPublication.bin').read_bytes());assert r.raw[:16]==b'NDMMATPUBLISH001'
    header=r.VectorTest();assert header[1]==revision and header[2]==E
    records=[];tag=r.VectorTest()
    while tag[0]>0:
        index=int(tag[0]);p,field,actual,before,after=[r.VectorTest() for _ in range(5)]
        item=dict(case=index,published=bool(tag[2]),uncertainty_available=bool(tag[4]),lambda_value=float(tag[3]),uncertainty=float(tag[5]))
        assert tag[7]==tag[8]==revision and tag[6]<=30
        if tag[2]:
            assert before[0]==after[0]==0 and tag[3]==pair['cases'][index-1]['lambda_value']
            np.testing.assert_array_equal(field,fields[index-1]);assert len(p)==60 and np.all(p[56:]==0)
            point,side=int(meta[0]),int(meta[1]);basis=op[side+1,point-1];native_gradient=E*p
            finite=np.zeros(60)
            for degree in range(56):
                coefficient=max(abs(basis[:,degree]))
                if coefficient==0:continue
                step=1e-9/coefficient;a=field.copy();b=field.copy();a[degree]+=step;b[degree]-=step
                finite[degree]=(env['TestYield'](point,side,a)-env['TestYield'](point,side,b))/(2*step)
            gradient_error=float(max(abs(finite-native_gradient))/max(abs(native_gradient)));assert gradient_error<1e-6
            item['independent_gradient_relative_error']=gradient_error
            if tag[4]:
                number=int(tag[10]);raw=(folder/f'Native/Adjoint.{number}.bin').read_bytes();assert raw[:16]==b'NDMADJOINTRES001'
                n,lam,tolerance=struct.unpack_from('<i2d',raw,16);assert n==61 and lam==tag[3] and len(raw)==36+8*(n*n+2*n)
                matrix=np.frombuffer(raw,dtype='<f8',count=n*n,offset=36).reshape(n,n)
                rhs=np.frombuffer(raw,dtype='<f8',count=n,offset=36+8*n*n)
                residual=np.frombuffer(raw,dtype='<f8',count=n,offset=36+8*(n*n+n))
                np.testing.assert_array_equal(rhs,np.eye(n)[-1]);np.testing.assert_allclose(matrix[:60,-1],native_gradient,rtol=1e-14,atol=1e-12)
                np.testing.assert_allclose(matrix[-1,:60],-(L@source['direction']),rtol=1e-14,atol=1e-12)
                weights_raw=(folder/f'Native/Adjoint.Weights.{number}.bin').read_bytes();assert struct.unpack_from('<i',weights_raw)[0]==n and len(weights_raw)==4+8*n
                weights=np.frombuffer(weights_raw,dtype='<f8',count=n,offset=4)
                backward=float(max(abs(matrix@weights-rhs)/np.maximum(abs(matrix)@abs(weights)+abs(rhs),np.finfo(float).tiny)))
                assert backward<512*n*np.finfo(float).eps
                columns=np.max(abs(matrix),axis=0);scaled=matrix/columns;rows=np.max(abs(scaled),axis=1)
                independent=np.linalg.solve(scaled/rows[:,None],rhs/rows)/columns
                weight_error=float(max(abs(independent-weights))/max(abs(weights)));assert weight_error<1e-6
                expected=float(meta[5]+meta[6]*max(abs((L@(source['offset']+lam*source['direction']))[:56])))
                assert abs(tolerance-expected)<1e-12*expected
                uncertainty=float(sum(abs(float(x)) for x in weights)*max(tolerance,float(max(abs(residual))))+2e-12*(1+abs(lam)))
                assert abs(uncertainty-tag[5])<=1e-10*max(uncertainty,1e-20)
                item.update(independent_uncertainty=uncertainty,adjoint_backward_error=backward,independent_weights_relative_error=weight_error)
        records.append(item);tag=r.VectorTest()
    assert tag[0]==0 and tag[2]==revision and r.position==len(r.raw)
    return records

def TestPair(folder,states,accepted,reference,revision):
    """Проверяет каждую пробу новой истории и истинный общий бюджет; возвращает фактические поля для проверки прогнозов."""
    env['states'].clear();env['states'].update({key:record['state'] for key,record in states.items()})
    r=R((folder/'Native/MaterialMidpointPair.bin').read_bytes());assert r.raw[:16]==b'NDMMATPAIR000001'
    meta,context,recorded,projection=[r.VectorTest() for _ in range(4)]
    np.testing.assert_array_equal(recorded,accepted);assert context.tolist()==[reference,float(revision)]
    assert meta[5:8].tolist()==[source['options'][3],source['options'][4],source['options'][2]]
    fields=[];factors=[];cases=[];steps=0;tag=r.VectorTest()
    while tag[0]>0:
        index=int(tag[0]);assert tag[2]==steps and tag[1]==(20 if index==1 else 30-steps)
        seed,initial,info=[r.VectorTest() for _ in range(3)];steps+=int(info[10]);assert steps<=30
        if index==2:
            np.testing.assert_array_equal(seed,.5*(accepted+fields[0]));assert tag[3]==.5*(reference+factors[0])
        item=dict(case=index,native_solved=bool(info[3]),steps=int(info[10]),lambda_value=float(info[6]),material_ready=bool(info[5]))
        if info[5]:
            field,values,forces=[r.VectorTest() for _ in range(3)];stress=TestStress(r)
            metrics,system=TestSystem(field,info[6],stress)
            assert max(abs(values[:56]-system['residual'][:56]))<2e-11
            assert max(abs(values[56:60]-E*system['residual'][56:]))<2e-11
            assert max(abs(forces-system['six_force_residual_vector']))<2e-11
            assert abs(values[-1]-E*(projection@field-meta[4]))<2e-11
            scalar=float(env['TestYield'](int(meta[0]),int(meta[1]),field)-info[1])
            difference=abs(scalar-info[7]);assert difference<2e-10
            assert abs(metrics['norm']-info[9])<2e-11
            item.update(metrics,scalar_residual=scalar,scalar_reconstruction_difference=float(difference),
                independently_equilibrated=bool(metrics['verified'] and abs(info[7])<=2e-12 and abs(scalar)<=2e-12+difference))
            fields.append(field);factors.append(float(info[6]))
        cases.append(item);tag=r.VectorTest()
    assert tag.tolist()==[0.,float(steps),float(revision)] and r.position==len(r.raw)
    return dict(cases=cases,steps=steps),fields,factors,meta

r=R((root/'Native/LoadingHistory.bin').read_bytes());assert r.raw[:16]==b'NDMLOADHIST00001'
stages=[];tag=r.VectorTest();previous_after=None
while tag[0]>0:
    assert tag[0]==1 and len(tag)==16;stage=int(tag[1]);assert stage==len(stages)+1
    folder=root if stage==1 else root/('Native/History%02d'%stage)
    revision,accepted,states=TestState(root/('Native/HistoryStateBefore.%d.bin'%stage))
    np.testing.assert_array_equal(accepted,trajectory[-1]);assert revision==tag[2] and tag[3]==lambdas[-1]
    if previous_after is not None:
        before=(root/('Native/HistoryStateBefore.%d.bin'%stage)).read_bytes()
        assert before==previous_after,'Between-step material history changed without acceptance'
    pair,fields,factors,meta=TestPair(folder,states,accepted,lambdas[-1],revision)
    publications=TestPublication(folder,pair,fields,meta,revision)
    if stage>1:
        fixture=R((folder/'Input/RecordedYieldSeed.bin').read_bytes());fixture.position=0
        fixture_meta,fixture_seed,fixture_accepted,fixture_expected,fixture_projection=[fixture.VectorTest() for _ in range(5)]
        assert fixture.position==len(fixture.raw)
        np.testing.assert_array_equal(fixture_accepted,accepted)
        scale=2 if growing and stages[-1]['total_steps_with_publications_and_post_commit']<=15 else 1
        np.testing.assert_array_equal(fixture_seed,trajectory[-1]+scale*(trajectory[-1]-trajectory[-2]))
        assert fixture_meta[3]==lambdas[-1]+scale*(lambdas[-1]-lambdas[-2])
        initial_increment*=scale
        accepted_value=env['TestYield'](int(meta[0]),int(meta[1]),accepted)
        assert abs((meta[2]-accepted_value)-initial_increment)<2e-12
        assert abs(meta[9]-(accepted_value+.5*initial_increment))<2e-12
    else:
        initial_increment=2*(meta[2]-meta[9])
    assert abs(TestH(states)-tag[10])<2e-12*max(abs(tag[10]),1e-20)
    item=dict(stage=stage,revision_before=revision,pair=pair,publications=publications,native_acceptance_gate=bool(tag[15]),H_before=float(tag[10]),
        H_full=float(tag[11]),H_resolution=float(tag[12]),signed_advance=float(tag[13]),accepted=False)
    item.update(native_field_length=float(tag[6]),native_field_error=float(tag[7]),native_lambda_pair_error=float(tag[8]),
        native_lambda_pair_bound=float(tag[9]),native_budget_used=int(tag[5]))
    if len(fields)==2:
        measured_length=TestArcLength(fields[0]-accepted)
        measured_error=TestArcLength(fields[0]-(2*fields[1]-accepted))/measured_length
        assert abs(measured_length-tag[6])<2e-12*measured_length and abs(measured_error-tag[7])<1e-10
        checks=dict(equilibria=all(c['independently_equilibrated'] for c in pair['cases']),remaining_budget=tag[5]<30,
            field_resolution=measured_length>16*pair['cases'][0]['bound']/E,field_pair_error=measured_error<=.005,
            lambda_pair_error=tag[8]<=tag[9],resolved_H_growth=tag[11]-tag[10]>tag[12],forward_direction=tag[13]>0)
        checks={key:bool(value) for key,value in checks.items()}
        assert bool(tag[15])==all(checks.values())
        item['independent_acceptance_checks']=checks
    if not tag[15]:
        stages.append(item);tag=r.VectorTest();break
    recorded,full,half=[r.VectorTest() for _ in range(3)]
    np.testing.assert_array_equal(recorded,accepted);np.testing.assert_array_equal(full,fields[0]);np.testing.assert_array_equal(half,fields[1])
    assert len(pair['cases'])==2 and all(c['independently_equilibrated'] for c in pair['cases']),(stage,pair)
    length=TestArcLength(full-accepted);error=TestArcLength(full-(2*half-accepted))/length
    assert abs(length-tag[6])<2e-12*length and abs(error-tag[7])<1e-10
    assert length>16*pair['cases'][0]['bound']/E and error<=.005 and tag[8]<=tag[9]
    assert tag[11]-tag[10]>tag[12] and tag[13]>0
    publication=r.VectorTest();assert publication[0]==2 and publication[1]==stage and publication[2]==1 and publication[3]==factors[0]
    actual=r.VectorTest();stress=TestStress(r);before_metrics,before_system=TestSystem(full,publication[3],stress);assert before_metrics['verified']
    assert max(abs((actual-before_system['actual_forces'])/before_system['force_scales']))<2e-11
    commit=r.VectorTest();assert commit[0]==3 and len(commit)==11 and commit[1]==stage and commit[2]==revision and commit[3]==revision+1
    after_revision,after_field,after_states=TestState(root/('Native/HistoryStateAfter.%d.bin'%stage))
    assert after_revision==revision+1;np.testing.assert_array_equal(after_field,full)
    maximum_stress_commit_error=0.;changed=0
    for key,old in states.items():
        new=after_states[key];point,side=key
        np.testing.assert_array_equal(new['state'][:2],old['state'][:2])
        assert np.all(new['damage'][:2]>=old['damage'][:2]-1e-16)
        changed+=int(not np.array_equal(new['state'],old['state']))
        maximum_stress_commit_error=max(maximum_stress_commit_error,float(max(abs(new['stress']-stress[side+1,point-1]))))
    assert maximum_stress_commit_error<2e-11 and changed>0
    assert commit[4]==commit[5]==tag[11] and commit[10]<=30
    assert abs(TestH(after_states)-commit[5])<2e-12*max(abs(commit[5]),1e-20)
    item.update(accepted=True,revision_after=after_revision,changed_concrete_states=changed,maximum_stress_commit_error=maximum_stress_commit_error,
        field_error=error,field_length=length,total_steps_with_publications_and_post_commit=int(commit[10]),
        post_commit_native_solved=bool(commit[6]),post_commit_native_exact=bool(commit[9]),post_commit_iterations=int(commit[8]),
        before_commit=before_metrics,H_after=float(commit[5]))
    if commit[6]:
        post_field,post_actual=[r.VectorTest() for _ in range(2)];post_stress=TestStress(r)
        metrics,post_system=TestSystem(post_field,commit[7],post_stress)
        assert max(abs((post_actual-post_system['actual_forces'])/post_system['force_scales']))<2e-11
        exact=bool(np.array_equal(post_field,full) and commit[7]==factors[0]);assert exact==bool(commit[9])
        item.update(after_commit=metrics,post_commit_fields_and_lambda_bitexact=exact,
            maximum_post_commit_stress_difference=float(np.max(abs(post_stress-stress))))
    trajectory.append(full);lambdas.append(factors[0]);previous_after=(root/('Native/HistoryStateAfter.%d.bin'%stage)).read_bytes()
    stages.append(item);tag=r.VectorTest()
assert tag.tolist()==[0.,float(sum(s['accepted'] for s in stages)),float(stages[-1].get('revision_after',stages[-1]['revision_before']))]
assert r.position==len(r.raw)
proof=dict(native=env['native'],actual_components=257,all256_numerical_components_exact980=True,original_first_pair_bitexact980=True,
    stages=stages,accepted_steps=sum(s['accepted'] for s in stages),
    all_accepted_states_idempotent=all(s.get('post_commit_fields_and_lambda_bitexact',False) and s.get('after_commit',{}).get('verified',False) and s.get('post_commit_iterations')==0 for s in stages if s['accepted']),
    no_future_equilibrium_input=True,original_materials_load_path_tolerances_and_shared30_per_step=True,
    automatic_cold_controller_not_tested=True,cold_gain_claimed=False,physical_capacity_certified=False,whole_goal_completed=False)
(root/'AcceptedLoadingHistoryProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for stage in stages:print('HISTORY STAGE',stage)
print('INDEPENDENT ACCEPTED HISTORY VERIFIED;ACCEPTED',proof['accepted_steps'],'IDEMPOTENT',proof['all_accepted_states_idempotent'])
