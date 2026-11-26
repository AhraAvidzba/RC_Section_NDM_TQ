"""ДЛЯ ТЕСТОВ. Независимо проверяет публикацию976 без движения, материальный градиент и фактическое сопряжённое решение."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json,sys,struct,csv,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,default=base/'Cold976PublishedMaterialPair_20261126')
root=parser.parse_args().root;parent=base/'Cold975ResolvedOutgoingPair_20261126'
sys.argv=[str(here/'Check-TestMidpointMaterialPair.py'),'--root',str(root)]
env=run_path(str(here/'Check-TestMidpointMaterialPair.py'),run_name='published_material_existing_pair_test')
R=env['R'];E=env['E'];source=env['source'];L=env['L'];yield_value=env['TestYield']
# Повторное чтение отделяет свидетельства полного отклика от дополнительных операций публикации.
r=R((root/'Native/MaterialMidpointPair.bin').read_bytes());meta,context,accepted,original_projection=[r.VectorTest() for _ in range(4)]
stresses=[];tag=r.VectorTest()
while tag[0]>0:
    seed,initial,info=[r.VectorTest() for _ in range(3)]
    if info[5]:
        field,values,forces=[r.VectorTest() for _ in range(3)];stress=np.zeros((3,43,6))
        for _ in range(129):
            point,side,code=r.VectorTest();stress[int(side)+1,int(point)-1]=r.VectorTest();assert code==0
        stresses.append(stress)
    tag=r.VectorTest()
assert r.position==len(r.raw)
same_binary=(root/'Native/MaterialMidpointPair.bin').read_bytes()==(parent/'Native/MaterialMidpointPair.bin').read_bytes()
r=R((root/'Native/MaterialPublication.bin').read_bytes());assert r.raw[:16]==b'NDMMATPUBLISH001'
header=r.VectorTest();assert header.tolist()==[context[0],context[1],E]

def TestDirectionMetrics(physical_field):
    """Измеряет направление любого доступного поля, отдельно от его равновесия и допустимости публикации."""
    yields={key:yield_value(*key,physical_field) for key in env['states']}
    beyond=[key for key,value in yields.items() if value>env['states'][key][8]*1e-11]
    old_strain=np.einsum('spcd,d->spc',env['op'],accepted-source['fields'][-2])
    new_strain=np.einsum('spcd,d->spc',env['op'],physical_field-accepted)
    old_length=env['TestLength'](accepted-source['fields'][-2])
    advance=float(np.einsum('spc,spc,p,c->',old_strain,new_strain,env['coords'][:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(env['coords'][:,2])*old_length))
    return dict(maximum_concrete_elastic_predictor_loading_value_MPa=max(yields.values()),
        concrete_predictors_positive_beyond_existing_return_tolerance=[list(key) for key in beyond],
        signed_spatial_advance_along_previous_secant=advance,previous_spatial_increment=old_length,
        forward_along_previous_spatial_secant=advance>0)

records=[];tag=r.VectorTest()
while tag[0]>0:
    index=int(tag[0]);assert index==len(records)+1 and len(tag)==11 and tag[7]==tag[8]==context[1]
    p,field,actual,before,after=[r.VectorTest() for _ in range(5)];case=env['cases'][index-1]
    item=dict(case=index,eligible=bool(tag[1]),published=bool(tag[2]),lambda_value=float(tag[3]),
        native_uncertainty_available=bool(tag[4]),native_uncertainty=float(tag[5]),total_steps=int(tag[6]),
        candidate_retained_after_uncertainty=bool(tag[9]),adjoint_index=int(tag[10]),
        before_effort=before.tolist(),after_effort=after.tolist())
    assert tag[6]<=30
    if case['material_ready']:
        item.update(TestDirectionMetrics(env['fields'][index-1]))
        item['direction_metrics_are_equilibrium']=case['independently_equilibrated']
    if tag[2]:
        assert before[0]==after[0]==0 and tag[3]==case['lambda_value']
        np.testing.assert_array_equal(field,env['fields'][index-1]);assert len(actual)==6 and len(p)==60
        assert np.all(p[56:]==0)
        point,side=int(meta[0]),int(meta[1]);basis=env['op'][side+1,point-1];native_gradient=E*p;differences=[]
        for strain_step in (1e-8,1e-9):
            finite=np.zeros(60)
            for degree in range(56):
                coefficient=max(abs(basis[:,degree]))
                if coefficient==0:continue
                h=strain_step/coefficient;plus=field.copy();minus=field.copy();plus[degree]+=h;minus[degree]-=h
                finite[degree]=(yield_value(point,side,plus)-yield_value(point,side,minus))/(2*h)
            relative=float(max(abs(finite-native_gradient))/max(abs(native_gradient)))
            differences.append(dict(physical_strain_difference_step=strain_step,maximum_relative_gradient_error=relative))
            assert relative<1e-6
        item.update(field_unchanged_bitexact=True,publication_iterations_zero=True,native_material_gradient_checked_by_independent_differences=differences)
        if tag[4]:
            adjoint_index=int(tag[10]);raw=(root/f'Native/Adjoint.{adjoint_index}.bin').read_bytes()
            assert raw[:16]==b'NDMADJOINTRES001'
            size,lam,tolerance=struct.unpack_from('<i2d',raw,16);assert size==61 and lam==tag[3]
            assert len(raw)==36+8*(size*size+2*size)
            A=np.frombuffer(raw,dtype='<f8',count=size*size,offset=36).reshape(size,size)
            rhs=np.frombuffer(raw,dtype='<f8',count=size,offset=36+8*size*size)
            residual=np.frombuffer(raw,dtype='<f8',count=size,offset=36+8*(size*size+size))
            np.testing.assert_array_equal(rhs,np.eye(size)[-1]);np.testing.assert_allclose(A[:60,-1],native_gradient,rtol=1e-14,atol=1e-12)
            np.testing.assert_allclose(A[-1,:60],-(L@source['direction']),rtol=1e-14,atol=1e-12);assert A[-1,-1]==0
            expected_tolerance=float(meta[5]+meta[6]*max(abs((L@(source['offset']+lam*source['direction']))[:56])))
            assert abs(tolerance-expected_tolerance)<1e-12*expected_tolerance
            weights_raw=(root/f'Native/Adjoint.Weights.{adjoint_index}.bin').read_bytes()
            assert struct.unpack_from('<i',weights_raw)[0]==size and len(weights_raw)==4+8*size
            weights=np.frombuffer(weights_raw,dtype='<f8',count=size,offset=4);assert np.all(np.isfinite(weights))
            backward=float(max(abs(A@weights-rhs)/np.maximum(abs(A)@abs(weights)+abs(rhs),np.finfo(float).tiny)))
            assert backward<512*size*np.finfo(float).eps
            columns=np.max(abs(A),axis=0);assert np.all(columns>0)
            scaled=A/columns;rows=np.max(abs(scaled),axis=1);assert np.all(rows>0)
            independent=np.linalg.solve(scaled/rows[:,None],rhs/rows)/columns
            relative=float(max(abs(independent-weights))/max(abs(weights)));assert relative<1e-6
            norm=max(tolerance,float(max(abs(residual))))
            uncertainty=float(sum(abs(float(x)) for x in weights)*norm+2e-12*(1+abs(lam)))
            assert abs(uncertainty-tag[5])<=1e-10*max(uncertainty,1e-20)
            outcomes=list(csv.reader((root/'Native/Adjoint.Outcomes.csv').read_text('utf-8-sig').splitlines(),delimiter=';'))
            outcome=next(x for x in outcomes if int(x[0])==adjoint_index);assert int(outcome[1])==1 and int(outcome[3])==0
            physical_scale=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
            force_norm=float(max(abs((actual-(source['offset']+lam*source['direction']))/physical_scale)))
            assert force_norm<=expected_tolerance+2e-11
            item.update(adjoint_native_backward_error=backward,adjoint_weights_relative_difference_from_independent_solve=relative,
                independent_uncertainty=uncertainty,adjoint_l1=float(sum(abs(float(x)) for x in weights)),
                uncertainty_norm_MPa=norm,actual_six_force_residual_MPa=force_norm,
                lambda_change_from_accepted=float(lam-context[0]),
                absolute_lambda_change_exceeds_uncertainty_and_noise=bool(abs(lam-context[0])>uncertainty+8*source['options'][4]*(1+abs(context[0]))))
    records.append(item);tag=r.VectorTest()
assert tag.tolist()==[0.,float(env['total']),float(context[1])] and r.position==len(r.raw)
assert len(records)==len(env['cases'])
proof=env['proof'];proof.update(publication_records=records,original_material_pair_binary_bitexact975=same_binary,
    eligible_publications=sum(x['eligible'] for x in records),successful_publications=sum(x['published'] for x in records),
    all_eligible_fields_published_unchanged=any(x['eligible'] for x in records) and all(not x['eligible'] or x['published'] and x.get('field_unchanged_bitexact',False) for x in records),
    all_publications_have_checked_native_lambda_uncertainty=any(x['published'] for x in records) and all(not x['published'] or x['native_uncertainty_available'] and 'independent_uncertainty' in x for x in records),
    native_lambda_uncertainty_not_recorded=False,no_material_history_acceptance=True,
    not_automatic_station_recovery_or_cold_search=True,physical_capacity_certified=False,cold_gain_claimed=False)
(root/'PublishedMaterialPairProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PUBLICATION AUDIT;PAIR BINARY EXACT975',same_binary,'RECORDS',records)
