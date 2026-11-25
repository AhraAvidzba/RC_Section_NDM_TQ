"""ДЛЯ ТЕСТОВ. Независимо проверяет принятую численную ветвь и половинные корни общей дуговой операции."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,default=base/'Cold823AutomaticLoadingWalk_20261125')
p.add_argument('--initial-input',type=Path);p.add_argument('--continued-prefix',action='store_true')
p.add_argument('--spectral-prefix',action='store_true');a=p.parse_args();root=a.root
assert not a.spectral_prefix or a.continued_prefix
parent=base/('Cold837ElasticPrincipalEvent_20261125' if a.spectral_prefix else 'Cold821PriorityLoadingArc_20261125')
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
names=['src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteResponse.cls',
    'src/Materials/CCoupledConcreteReturn.cls','Input/SourceColdRawHistory.bin']
if a.spectral_prefix:
    from hashlib import sha256
    prep=json.loads((root/'Preparation.json').read_bytes());audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
    assert audit['components']==256 and audit['imported_this_pass']==['CCoupledArcContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']
    assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==253
    for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls'):
        assert sha256((root/name).read_bytes()).hexdigest().upper()==prep['inputs'][name]
    names+=['src/Materials/CCoupledReinforcedResponse.cls','src/Solver/CCoupledMaterialConstraint.cls']
else:names+=['src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls']
for name in names:
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
assert (root/'Input/Reverse1.bin').read_bytes()==(a.initial_input or parent/'Input/Reverse1.bin').read_bytes()
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'LoadingWalk.Accepted.Sequence')
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
n=len(source['fields']);assert n==(86 if a.spectral_prefix else (69 if a.continued_prefix else 43)) and len(fields)>n
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
proof=dict(native=native,all_working_methods_and_materials_exact821=not a.spectral_prefix,
    spectral_event_integration=a.spectral_prefix,material_and_scalar_constraint_exact837=a.spectral_prefix,
    initial_source_states=n,new_accepted_equilibria=accepted,
    continued_from_verified69_prefix=a.continued_prefix and not a.spectral_prefix,
    continued_from_verified86_prefix=a.spectral_prefix,
    raw_force_and_moment_comparison_uses_actual_equilibrium_scales=True,raw_force_equilibrium_tolerance_not_weakened=True,
    total_calls=calls,attempts=proofs,all_equilibria=mechanics,all_half_roots_independently_rebuilt=True,
    maximum_three_station_steel_utilization=float(np.max(utilization)),all_steel_strain_limits_verified_at_all_three_stations=True,
    actual_material_parameters_unchanged=True,no_future_field_used=True,
    physical_admissibility_ONQ012_open=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'AutomaticLoadingWalkProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root.name,'NEW EQUILIBRIA',accepted,'CALLS',calls,'LAST LAMBDA',mechanics['states'][-1]['lambda'],
    'ALL HALF ROOTS/SIX FORCES/THREE-STATION STEEL VERIFIED; NO PHYSICAL CAPACITY')
