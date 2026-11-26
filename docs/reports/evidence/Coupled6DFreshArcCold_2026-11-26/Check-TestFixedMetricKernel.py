"""ДЛЯ ТЕСТОВ. Независимо восстанавливает силы, кривизну, ориентацию и сталь автоматической ветви общего ядра с постоянной метрикой."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,default=base/'Cold867FixedMetricKernel_20261126');args=p.parse_args();root=args.root
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and audit['imported_this_pass']==['CCoupledArcContinuation','CCoupledDeformationContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==252
scope=json.loads((root/'Preparation.json').read_bytes())['directed_scope'];assert scope['case']=='GeneralFixedMetricKernelFromCold19'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native['failed']==0 and not native['timeout600'] and native['own_process_closed']
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];source=read(root/'Input/Reverse1.bin')
assert len(source['fields'])==19
assert (root/'Input/Reverse1.bin').read_bytes()==(base/'Cold866Cold19ReferenceMetricPairs_20261126/Input/Reverse1.bin').read_bytes()
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'LoadingWalk.Accepted.Sequence')
n=len(source['fields']);np.testing.assert_allclose(fields[:n],source['fields'],rtol=1e-14,atol=1e-300)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/LoadingWalk.Accepted.Sequence.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu));D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
net=np.tile(D,(43,1,1));net[25:]=200000*np.einsum('pc,pd->pcd',axis[25:],axis[25:])-D
elastic=g+np.einsum('pcr,pcd,pds->rs',w,net,B)+np.einsum('pcr,pcd,pds->rs',v,net,gradient)
linear=np.linalg.solve(elastic,L@source['direction'])

def TestLength(delta):
    """Исключает реакции связей из физической тензорной метрики трёх станций."""
    return float(np.sqrt(max(0.,delta@M@delta)))

alpha=TestLength(linear);assert alpha>0
force_scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])

def TestBound(factor):
    """Применяет неизменную строгую норму равновесия к фактической нагрузке, без подмены размерностей."""
    return float(.0001*(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+factor*source['direction']))[:56]))))

def TestH(phase,count):
    """Суммирует энергетические координаты бетона, которые не являются повреждением или физическим пределом."""
    total=0.;seen=set();gt,gc,ell=source['calibration'][6:9]
    for row in validate['TestRows'](root/'Native'/(phase+'.csv')):
        if row[0]=='S' and int(row[1])==count and int(row[3])<=25:
            key=(int(row[2]),int(row[3]));assert key not in seen;seen.add(key)
            total+=coords[key[1]-1,2]*(gt*float(row[10])+gc*float(row[11]))/ell
    assert len(seen)==75
    return total/(3*sum(coords[:25,2]))

def TestSteel(field):
    """Проверяет продольную и поперечную сталь с разными предельными деформациями на всех трёх станциях."""
    strain=np.einsum('spcd,d->spc',op,field)
    axial=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
    limit=np.where(long[None,:],np.where(axial>=0,.025,.018),.02)
    util=float(np.max(abs(axial)/limit));assert util<=1
    return util

Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
rows=validate['TestRows'](root/'Native/LoadingWalk.Attempts.csv');assert 1<=len(rows)<=32
raw_fields=list(source['fields']);factors=list(source['lambdas']);accepted=calls=attempts_at_revision=0;proofs=[]
old_tangent=json.loads((base/'Cold864Cold19BranchTangents_20261125/ColdBranchTangentsProof.json').read_bytes())['tangents'][1]
previous_ratio=None;previous_solved=False;previous_limit=None;previous_step_error=None
for row in rows:
    attempt,count,solved,points,ratio,factor,before,after,used,total,primary,error,kind=row
    attempt,count,points,before,after,used,total,kind=map(int,(attempt,count,points,before,after,used,total,kind))
    solved=solved=='True';ratio,factor,primary,error=map(float,(ratio,factor,primary,error))
    assert count==n+accepted and attempts_at_revision<6
    if previous_ratio is not None:
        expected=min(previous_limit,.8*np.sqrt(.005/max(previous_step_error,1e-8))) if previous_solved else previous_ratio*.5
        assert abs(ratio-expected)<=1e-13*max(1.,abs(expected))
    attempts_at_revision+=1;calls+=used;assert total==calls<=96
    r=Reader((root/'Native'/('LoadingWalk.'+str(attempt).zfill(3)+'.bin')).read_bytes());assert r.raw[:16]==b'NDMOPERARC000004'
    packet=[r.VectorTest() for _ in range(15)];info=packet[0];assert len(info)==34
    returned=actual=projection=None
    if info[3]:returned,actual,projection=[r.VectorTest() for _ in range(3)]
    effort=r.VectorTest();half=r.VectorTest();half_factor=r.VectorTest();assert r.position==len(r.raw) and len(half_factor)==1
    assert info[14]==info[15]==before and info[16]<=30 and info[17]<=30 and info[18]==used and info[19]==primary
    assert info[28]==kind and bool(info[3])==solved
    assert abs(info[9]-alpha)<2e-13*alpha and abs(info[31]-alpha)<2e-13*alpha
    np.testing.assert_array_equal(packet[2],raw_fields[-1]);reference=raw_fields[-1];reference_lambda=factors[-1]
    guide=Reader((root/'Native'/('LoadingWalk.Guide'+str(attempt).zfill(3)+'.bin')).read_bytes());assert guide.raw[:16]==b'NDMWALKH00000003'
    gm=guide.VectorTest();assert gm[0]==attempt and gm[1] in (1,2)
    guides=[]
    for size in (3,3,4):
        present=guide.VectorTest()[0];assert present in (0,1)
        guides.append([guide.VectorTest() for _ in range(size)] if present else None)
    assert guide.position==len(guide.raw)
    tangent=guides[2];assert bool(info[32])==(tangent is not None)
    item=dict(attempt=attempt,source_count=count,solved=solved,kind=kind,ratio=ratio,calls=used,total_calls=calls,primary=int(primary),source_revision_unchanged=True,tangent_used=tangent is not None)
    if tangent is not None:
        tm,probe,perlambda,movement=tangent
        np.testing.assert_array_equal(probe,reference+(reference-raw_fields[-2])/1024.)
        np.testing.assert_allclose(movement,tm[1]*perlambda,rtol=1e-14,atol=1e-300)
        assert abs(TestLength(perlambda)-tm[2])<1e-12*tm[2] and abs(TestLength(movement)-TestLength(reference-raw_fields[-2]))<1e-12*TestLength(movement)
        orient=float((reference-raw_fields[-2])@M@perlambda+alpha*alpha*(reference_lambda-factors[-2]))
        assert abs(orient-tm[3])<1e-10*abs(orient) and np.sign(tm[1])==np.sign(orient)
        expected_arc=float(np.sqrt(TestLength(movement)**2+alpha**2*tm[1]**2))
        np.testing.assert_allclose(packet[1],M@movement/expected_arc,rtol=1e-8,atol=1e-17)
        assert abs(info[12]-alpha**2*tm[1]/expected_arc)<1e-18
        if count==19:
            np.testing.assert_allclose(movement,np.array(old_tangent['same_length_field_movement']),rtol=2e-7,atol=1e-14)
            assert abs(tm[1]-old_tangent['same_length_lambda_increment'])<2e-7*abs(tm[1])
        item['signed_tangent_lambda_increment']=float(tm[1])
    if solved:
        assert info[2] and effort[0]==0 and after==before+1 and kind in (1,2,3,4)
        factor=float(info[23]);half_lambda=float(half_factor[0]);assert factor>0 and factor<=source['options'][0]
        length=TestLength(returned-reference);assert length>0
        field_defect=TestLength(returned-(2*half-reference));field_error=field_defect/length
        le=abs(2*half_lambda-reference_lambda-factor)
        distance=float(np.sqrt(length**2+alpha**2*(factor-reference_lambda)**2))
        joint=float(np.sqrt(field_defect**2+alpha**2*le**2)/distance)
        noise=8*.0001*source['options'][4]*(1+abs(reference_lambda))
        allowance=max(noise,np.sqrt(max(0.,(.005*distance)**2-field_defect**2))/alpha)
        assert field_error<=.005 and le<=allowance
        step_error=max(field_error,float(np.sqrt(field_defect**2+alpha**2*(0 if le<=noise else le)**2)/distance))
        np.testing.assert_allclose(gm[2],step_error,rtol=1e-5,atol=1e-10)
        assert .5*length*E>8*max(TestBound(reference_lambda),TestBound(factor),TestBound(half_lambda))
        cosine=float((reference-raw_fields[-2])@M@(returned-reference)/(TestLength(reference-raw_fields[-2])*length));assert cosine>0
        if kind==1:
            np.testing.assert_array_equal(half,packet[5]);assert half_lambda==info[5]
            np.testing.assert_array_equal(returned,packet[6])
            assert returned.tobytes()==packet[6].tobytes()
            np.testing.assert_allclose(info[20],field_error,rtol=1e-5,atol=1e-10)
            np.testing.assert_allclose(info[30],joint,rtol=1e-5,atol=1e-10)
            for point,lam,fraction in ((returned,factor,1.),(half,half_lambda,.5)):
                progress=float(packet[1]@(point-reference)+info[12]*(lam-reference_lambda))
                assert progress>0 and np.sqrt(TestLength(point-reference)**2+alpha**2*(lam-reference_lambda)**2)<=3*fraction*info[11]
                assert abs(progress-fraction*info[11])*E<=TestBound(lam)+1e-12
        accepted+=1;assert points==accepted
        full_weak=mechanics['states'][n+accepted-1]
        assert full_weak['residual_MPa']<=TestBound(factor)+1e-12 and full_weak['force_residual_MPa']<=TestBound(factor)+1e-12
        np.testing.assert_allclose(fields[n+accepted-1],returned,rtol=1e-14,atol=1e-300)
        phase='LoadingWalk.Half'+str(accepted).zfill(3)+'.Sequence';hp,hf=validate['TestValidate'](root,phase)
        assert len(hf)==count+1
        np.testing.assert_allclose(hf[:-1],raw_fields,rtol=1e-14,atol=1e-300);np.testing.assert_allclose(hf[-1],half,rtol=1e-14,atol=1e-300)
        hw=hp['states'][-1];assert hw['residual_MPa']<=TestBound(half_lambda)+1e-12 and hw['force_residual_MPa']<=TestBound(half_lambda)+1e-12
        raw_force=float(np.max(abs((actual-source['offset']-factor*source['direction'])/force_scales)));assert raw_force<=TestBound(factor)
        full_util=TestSteel(returned);half_util=TestSteel(half)
        h0=TestH('LoadingWalk.Accepted.Sequence',count);hfull=TestH('LoadingWalk.Accepted.Sequence',count+1);hhalf=TestH(phase,count+1)
        assert hfull>h0 and hhalf>h0
        item.update(lambda_value=factor,half_lambda=half_lambda,field_error=field_error,joint_error=joint,step_error=step_error,lambda_error=le,lambda_bound=float(allowance),
            weak=full_weak,half_weak=hw,raw_force_residual_MPa=raw_force,steel_util=full_util,half_steel_util=half_util,H_increment=hfull-h0,H_half_increment=hhalf-h0,direction_cosine=cosine,
            restoration_zero_iterations_same_lambda_same_field_native_guard=True)
        raw_fields.append(returned);factors.append(factor);attempts_at_revision=0
    else:assert after==before and points==accepted
    previous_ratio=ratio;previous_solved=solved;previous_limit=float(gm[1]);previous_step_error=float(gm[2]);proofs.append(item)
assert accepted==len(fields)-n
proof=dict(native=native,actual_components=256,working_kernel_changed_and_verified=True,source19_true_history=True,no_future_field_or_artificial_history=True,
    independent_elastic_alpha=alpha,initial_states=n,new_accepted_equilibria=accepted,total_calls=calls,attempts=proofs,
    independently_rebuilt_half_root_count=accepted,all_six_forces_and_56physical_plus4gauges_checked=True,
    unchanged_materials_and_force_tolerances=True,all_new_H_strictly_growing=accepted>0,
    local_curve_control_not_sequential_half_step_integration_error=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'FixedMetricKernelProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('GENERAL KERNEL:ACCEPTED',accepted,'CALLS',calls,'LAST LAMBDA',factors[-1],';ALL RETURNED/HALF FORCES,STEEL,METRIC,H CHECKED;NO PHYSICAL LIMIT')
