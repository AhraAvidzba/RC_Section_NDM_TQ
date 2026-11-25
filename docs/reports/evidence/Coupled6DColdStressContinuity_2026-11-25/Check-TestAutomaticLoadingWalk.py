"""ДЛЯ ТЕСТОВ. Независимо проверяет принятую численную ветвь и половинные корни общей дуговой операции."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,default=base/'Cold823AutomaticLoadingWalk_20261125')
p.add_argument('--initial-input',type=Path);p.add_argument('--continued-prefix',action='store_true')
p.add_argument('--spectral-prefix',action='store_true');p.add_argument('--h-reentry-prefix',action='store_true');p.add_argument('--auto-h-guide-prefix',action='store_true');p.add_argument('--loading-secant-prefix',action='store_true');p.add_argument('--cold19-restart',action='store_true');a=p.parse_args();root=a.root
assert not a.spectral_prefix or a.continued_prefix
assert not (a.spectral_prefix and a.h_reentry_prefix)
assert not a.auto_h_guide_prefix or (a.continued_prefix and not a.spectral_prefix and not a.h_reentry_prefix)
assert not a.loading_secant_prefix or a.auto_h_guide_prefix
assert not a.cold19_restart or a.loading_secant_prefix
parent=base/('Cold851LoadingSecantReparameterization_20261125' if a.loading_secant_prefix else ('Cold845SteelActivityArcPolicy_20261125' if a.auto_h_guide_prefix else ('Cold841HGuideMaterialEvent_20261125' if a.h_reentry_prefix else ('Cold837ElasticPrincipalEvent_20261125' if a.spectral_prefix else 'Cold821PriorityLoadingArc_20261125'))))
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
names=['src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteResponse.cls',
    'src/Materials/CCoupledConcreteReturn.cls','Input/SourceColdRawHistory.bin']
if a.cold19_restart:names.remove('Input/SourceColdRawHistory.bin')
if a.spectral_prefix:
    from hashlib import sha256
    prep=json.loads((root/'Preparation.json').read_bytes());audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
    assert audit['components']==256 and audit['imported_this_pass']==['CCoupledArcContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']
    assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==253
    for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls'):
        assert sha256((root/name).read_bytes()).hexdigest().upper()==prep['inputs'][name]
    names+=['src/Materials/CCoupledReinforcedResponse.cls','src/Solver/CCoupledMaterialConstraint.cls']
else:names+=['src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls']
if a.h_reentry_prefix or a.auto_h_guide_prefix:
    names+=['src/Materials/CCoupledReinforcedResponse.cls','src/Solver/CCoupledMaterialConstraint.cls']
    audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
    assert audit['components']==256 and audit['imported_this_pass']==['modTestCoupledElasticState']
    assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==255
for name in names:
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
if a.cold19_restart:
    prep=json.loads((root/'Preparation.json').read_bytes());scope=prep['directed_scope']
    assert scope['case']=='AutomaticArcRestartFromVerifiedCold19' and scope['actual_imports_this_replay']==[]
    compiled=root.parent/scope['compiled_source_root'];cold=root.parent/scope['source_root'];verified=root.parent/scope['verified_stress_root']
    assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(compiled/'docs/regression/Diagnostic.xlsm').read_bytes()
    assert (root/'tests/modTestCoupledElasticState.bas').read_bytes()==(compiled/'tests/modTestCoupledElasticState.bas').read_bytes()
    assert (root/'Input/SourceColdRawHistory.bin').read_bytes()==(cold/'Native/Stock.RawHistory.bin').read_bytes()
    assert (root/'Input/Reverse1.bin').read_bytes()==(verified/'Input/Reverse1.bin').read_bytes()
    verified_proof=json.loads((verified/'VerifiedColdStressReplayProof.json').read_bytes())
    assert verified_proof['source_states']==19 and verified_proof['source_workflow_incomplete'] and verified_proof['all_raw_fields_fractions_and_lambdas_exact']
else:assert (root/'Input/Reverse1.bin').read_bytes()==(a.initial_input or parent/'Input/Reverse1.bin').read_bytes()
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'LoadingWalk.Accepted.Sequence')
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
n=len(source['fields']);assert n==(19 if a.cold19_restart else (70 if a.auto_h_guide_prefix else (71 if a.h_reentry_prefix else (86 if a.spectral_prefix else (69 if a.continued_prefix else 43)))))
assert len(fields)>=n if a.cold19_restart else len(fields)>n
np.testing.assert_allclose(fields[:n],source['fields'],rtol=1e-14,atol=1e-300)
ops=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/LoadingWalk.Accepted.Sequence.Operators.csv',60,43)
g,L,w,v,coords,axis,b,gradient=ops;E=source['calibration'][0]
force_scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
op=np.array([b+s*.001*gradient for s in (-1,0,1)])
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1.,1.,1.,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0

def TestLength(delta):
    """Измеряет физическое тензорное приращение трёх станций без реакций связей."""
    return float(np.sqrt(max(0.,delta@M@delta)))

def TestBound(factor):
    """Восстанавливает строгий внутренний допуск на фактической нагрузке All6."""
    return float(.0001*(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+factor*source['direction']))[:56]))))

Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
rows=validate['TestRows'](root/'Native/LoadingWalk.Attempts.csv');assert 1<=len(rows)<=32
accepted=0;calls=0;attempts_at_revision=0;proofs=[];raw_fields=list(source['fields'])
for row in rows:
    attempt,count,solved,points,ratio,factor,before,after,used,total,primary,error,kind=row
    attempt,count,points,before,after,used,total,kind=map(int,(attempt,count,points,before,after,used,total,kind))
    solved=solved=='True';ratio,factor,primary,error=map(float,(ratio,factor,primary,error))
    assert count==n+accepted and attempts_at_revision<4
    attempts_at_revision+=1;calls+=used;assert calls==total<=96
    r=Reader((root/'Native'/('LoadingWalk.'+str(attempt).zfill(3)+'.bin')).read_bytes())
    signature=r.raw[:16];assert signature in (b'NDMOPERARC000003',b'NDMOPERARC000004')
    packet=[r.VectorTest() for _ in range(15)];info=packet[0];returned=actual=projection=None
    if info[3]:returned,actual,projection=[r.VectorTest() for _ in range(3)]
    effort=r.VectorTest();selected_half=selected_half_lambda=None
    if signature==b'NDMOPERARC000004':
        selected_half=r.VectorTest();selected_factor=r.VectorTest();assert len(selected_factor)==1
        selected_half_lambda=float(selected_factor[0])
    assert r.position==len(r.raw)
    assert info[14]==info[15]==before and info[16]<=30 and info[17]<=30 and info[18]==used
    assert info[19]==primary and kind==info[28] and bool(info[3])==solved
    np.testing.assert_array_equal(packet[2],raw_fields[-1])
    assert info[19]==(info[16] if kind==1 else info[16]+info[17])
    item=dict(attempt=attempt,source_count=count,solved=solved,calls=used,total_calls=total,primary=int(primary),kind=kind,
        revision_before=before,revision_after_probe=int(info[15]),revision_after_caller=after)
    if solved:
        assert info[2] and effort[0]==0 and after==before+1 and kind in (1,2,3,4)
        assert abs(info[23]-factor)<=5e-15
        factor=float(info[23])  # Все расчётные проверки используют исходный Double; CSV нужен только для читаемого журнала.
        assert factor>0 and factor<=source['options'][0]
        if kind==1:half,half_lambda=packet[5],info[5]
        elif kind in (2,4):half,half_lambda=packet[11],info[24]
        else:
            assert signature==b'NDMOPERARC000004'
            half,half_lambda=selected_half,selected_half_lambda
        if signature==b'NDMOPERARC000004':
            assert np.array_equal(selected_half,half) and selected_half_lambda==half_lambda
        reference=raw_fields[-1];reference_lambda=info[6]
        length=TestLength(returned-reference);assert length>0
        field_error=TestLength(returned-(2*half-reference))/length
        le=abs(2*half_lambda-reference_lambda-factor)
        lb=max(.005*abs(factor-reference_lambda),8*.0001*source['options'][4]*(1+abs(reference_lambda)))
        np.testing.assert_allclose(error,field_error,rtol=1e-4,atol=1e-11)
        assert field_error<=.005 and le<=lb
        assert .5*length*E>8*max(TestBound(reference_lambda),TestBound(factor),TestBound(half_lambda))
        previous=raw_fields[-2];cos=float((reference-previous)@M@(returned-reference)/(TestLength(reference-previous)*length))
        assert cos>0
        accepted+=1;assert points==accepted
        weak=mechanics['states'][n+accepted-1]
        assert weak['residual_MPa']<=TestBound(factor)+1e-12 and weak['force_residual_MPa']<=TestBound(factor)+1e-12
        np.testing.assert_allclose(fields[n+accepted-1],returned,rtol=1e-14,atol=1e-300)
        hp,hf=validate['TestValidate'](root,'LoadingWalk.Half'+str(accepted).zfill(3)+'.Sequence')
        assert len(hf)==count+1
        np.testing.assert_allclose(hf[:-1],raw_fields,rtol=1e-14,atol=1e-300)
        np.testing.assert_allclose(hf[-1],half,rtol=1e-14,atol=1e-300)
        hlast=hp['states'][-1]
        assert hlast['residual_MPa']<=TestBound(half_lambda)+1e-12 and hlast['force_residual_MPa']<=TestBound(half_lambda)+1e-12
        # Силы и моменты имеют разные единицы: используется та же физическая норма МПа и тот же внутренний допуск, что в независимом равновесии.
        raw_force_residual=float(np.max(abs((actual-source['offset']-factor*source['direction'])/force_scales)))
        assert raw_force_residual<=TestBound(factor)
        item.update(lambda_value=factor,half_lambda=float(half_lambda),field_half_error=field_error,lambda_half_error=float(le),
            lambda_half_bound=float(lb),mechanical_direction_cosine=cos,weak=weak,half_weak=hlast,
            raw_six_force_residual_MPa=raw_force_residual)
        raw_fields.append(returned);attempts_at_revision=0
    else:
        assert after==before and points==accepted
    proofs.append(item)
assert accepted==len(fields)-n
guards=(root/'Native/RecentLoading.Guards.txt').read_text('cp1251').splitlines()
expected_first=('NO_RECENT_BOUNDARY_EXPECTED;ARC_SECANT_USED;REVISION_UNCHANGED' if a.continued_prefix else
    'ZERO_REJECTED;UNLOADING_REJECTED;FORWARD_SELECTED;REVISION_UNCHANGED')
assert guards==[expected_first,'RESET_CONTEXT_REJECTED;REVISION_ADVANCED']
driver=(root/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
assert 'TestInputCell(steel, "Steel.TwoLine.Es2").Value2 = 0.018' in driver
assert 'TestInputCell(steel, "Steel.TwoLine.Es2", 1, 3).Value2 = 0.025' in driver
assert 'grades.Cells(2, 6).Value2 = 0.02: grades.Cells(2, 7).Value2 = 0.02' in driver
strains=np.einsum('spcd,nd->nspc',op,fields)
axial=np.einsum('nspc,pc->nsp',strains[:,:,25:],axis[25:])
longitudinal=axis[25:,2]==1;assert set(axis[25:,2])=={0.,1.}
limits=np.where(longitudinal[None,None,:],np.where(axial>=0,.025,.018),.02)
utilization=abs(axial)/limits;assert np.max(utilization)<=1
proof=dict(native=native,all_working_methods_and_materials_exact821=not (a.spectral_prefix or a.h_reentry_prefix),
    spectral_event_integration=a.spectral_prefix or a.h_reentry_prefix,material_and_scalar_constraint_exact837=a.spectral_prefix or a.h_reentry_prefix,
    automatic_departure_from_verified_H_face=a.h_reentry_prefix,
    initial_source_states=n,new_accepted_equilibria=accepted,
    continued_from_verified69_prefix=a.continued_prefix and not a.spectral_prefix,
    continued_from_verified86_prefix=a.spectral_prefix,
    raw_force_and_moment_comparison_uses_actual_equilibrium_scales=True,raw_force_equilibrium_tolerance_not_weakened=True,
    total_calls=calls,attempts=proofs,all_equilibria=mechanics,all_half_roots_independently_rebuilt=accepted>0,
    independently_rebuilt_half_root_count=accepted,
    maximum_three_station_steel_utilization=float(np.max(utilization)),all_steel_strain_limits_verified_at_all_three_stations=True,
    actual_material_parameters_unchanged=True,no_future_field_used=True,
    physical_admissibility_ONQ012_open=True,physical_capacity_certified=False,whole_goal_completed=False)
if a.auto_h_guide_prefix:
    proof['all_working_methods_and_materials_exact821']=False
    proof['all_working_methods_and_materials_exact851' if a.loading_secant_prefix else 'all_working_methods_and_materials_exact845']=True
    proof['continued_from_verified69_prefix']=False
    proof['automatic_Hguide_from_exact70_without_prescribed_face']=not a.cold19_restart
    proof['restart_from_verified_cold19_not_new_cold_capacity']=a.cold19_restart
    previous_limit=None;previous_solved=False;guide_attempts=[];loading_guide_attempts=[];loading_cases={}
    for row,item in zip(rows,proofs):
        attempt=item['attempt'];r=Reader((root/'Native'/('LoadingWalk.Guide'+str(attempt).zfill(3)+'.bin')).read_bytes())
        assert r.raw[:16]==(b'NDMWALKH00000002' if a.loading_secant_prefix else b'NDMWALKH00000001')
        meta=r.VectorTest();assert meta[0]==attempt and meta[1] in (1,2)
        guide=[r.VectorTest() for _ in range(3)] if meta[2] else None
        loading=None
        if a.loading_secant_prefix:
            present=r.VectorTest()[0];assert present in (0,1)
            if present:loading=[r.VectorTest() for _ in range(3)]
        assert r.position==len(r.raw)
        if previous_solved:assert float(row[4])<=previous_limit
        if guide is not None:
            assert len(guide[0])==17 and item['revision_before']==guide[0][14]==guide[0][15]
            assert guide[0][1]>=1 and guide[0][2]<=15
            guide_attempts.append(dict(attempt=attempt,meta=guide[0].tolist()))
        if loading is not None:
            loading_meta,loading_field,prediction=loading;reference=raw_fields[item['source_count']-1]
            assert loading_meta[5]==item['revision_before'] and loading_meta[6]==0
            distance=TestLength(loading_field-reference);predicted_distance=TestLength(prediction-reference)
            assert abs(distance-loading_meta[2])<=1e-18 and abs(predicted_distance-loading_meta[3])<=1e-18
            assert loading_meta[4]==loading_meta[3]/loading_meta[2]
            expected=reference.copy();expected[:56]+=loading_meta[4]*(loading_field[:56]-reference[:56])
            np.testing.assert_array_equal(expected,prediction)
            loading_guide_attempts.append(dict(attempt=attempt,meta=loading_meta.tolist(),guide_not_accepted=True))
            loading_cases[1000+attempt]=dict(field=loading_field,lambda_value=float(loading_meta[1]),accuracy=.0001)
        previous_limit=float(meta[1]);previous_solved=item['solved']
    if not a.cold19_restart:assert guide_attempts and proofs[0]['solved'] and proofs[0]['kind']==2 and guide_attempts[0]['attempt']==1
    if a.cold19_restart:
        for key in ('calibration','options','offset','direction'):
            np.testing.assert_array_equal(source[key],run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](verified/'Input/Reverse1.bin')[key])
    elif a.loading_secant_prefix:
        parent_input=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
        assert len(parent_input['fields'])==86 and len(source['fields'])==70
        for key in ('calibration','options','offset','direction'):np.testing.assert_array_equal(source[key],parent_input[key])
        for key in ('fields','fractions','lambdas'):np.testing.assert_array_equal(source[key],parent_input[key][:70])
    else:assert (root/'Input/Reverse1.bin').read_bytes()==(parent/'Input/Reverse1.bin').read_bytes()
    assert not (root/'Input/HMaterialGuide.bin').exists()
    energies=np.zeros((len(fields),3,25,2));seen=np.zeros((len(fields),3,25),int)
    for row in validate['TestRows'](root/'Native/LoadingWalk.Accepted.Sequence.csv'):
        if row[0]=='S' and int(row[3])<=25:
            frame,side,point=int(row[1])-1,int(row[2])+1,int(row[3])-1
            energies[frame,side,point]=np.array(row[10:12],float);seen[frame,side,point]+=1
    assert np.all(seen==1)
    gt,gc,localization=source['calibration'][6:9]
    h=np.einsum('fspc,p,c->f',energies,coords[:25,2],np.array([gt/localization,gc/localization]))/(3*sum(coords[:25,2]))
    proof['automatic_Hguide_attempts']=guide_attempts
    if a.loading_secant_prefix:
        if not a.cold19_restart:assert loading_guide_attempts
        loading_mechanics={}
        if loading_cases:loading_mechanics,_=run_path(str(here/'TestUnacceptedMechanics.py'))['TestUnacceptedRoots'](root,source,loading_cases,phase='LoadingWalk.Accepted.Sequence')
        proof['loading_guide_attempts']=loading_guide_attempts
        proof['loading_guide_independent_mechanics']=loading_mechanics
    proof['accepted_H_MPa']=h.tolist()
    proof['new_H_increment_MPa']=np.diff(h[n-1:]).tolist()
    proof['all_new_H_strictly_growing']=bool(accepted>0 and np.all(np.diff(h[n-1:])>0))
(root/'AutomaticLoadingWalkProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root.name,'NEW EQUILIBRIA',accepted,'CALLS',calls,'LAST LAMBDA',mechanics['states'][-1]['lambda'],
    ('ALL HALF ROOTS/SIX FORCES/THREE-STATION STEEL VERIFIED; NO PHYSICAL CAPACITY' if accepted else
     'NO NEW OR HALF ROOTS ACCEPTED;INPUT SIX FORCES/STEEL VERIFIED;NO PHYSICAL CAPACITY'))
