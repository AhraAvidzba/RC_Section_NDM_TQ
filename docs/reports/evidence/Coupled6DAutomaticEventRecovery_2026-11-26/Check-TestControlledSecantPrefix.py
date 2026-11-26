"""ДЛЯ ТЕСТОВ. Независимо собирает равновесие и половинную метрику префикса исходной меры без будущего корня."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,default=base/'Cold965ControlledSecantPrefix_20261126')
parser.add_argument('--root-budget',type=int,choices=[4,10],default=4)
args=parser.parse_args();root=args.root;parent=base/'Cold960ResolvedFieldProgress_20261126'
native,native_text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
# Первый вход965 напечатал собственный итог в другом формате; он проверяется явно, но не прибавляется к счётчику оболочки.
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
assert 'TOTAL_CONTROLLED_SECANT_PREFIX_TEST;1;0;root_and_half_check_not_capacity' in native_text
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
imports=['CCoupledNewtonSolver','modTestCoupledElasticState'] if args.root_budget==4 else ['CCoupledNewtonSolver']
assert audit['components']==257 and audit['imported_this_pass']==imports
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==257-len(imports)
before=(parent/'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
after=(root/'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
block=(root/'TestTryControlledSecantPrefix.cls.txt').read_text('utf-8-sig').rstrip()
if args.root_budget!=4:block=block.replace('If budget > 4 Then budget = 4','If budget > 10 Then budget = 10',1)
assert after==before+'\n\n'+block+'\n'
assert (root/'Input/Reverse1.bin').read_bytes()==(parent/'Input/Reverse1.bin').read_bytes()
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
validate=run_path(str(here/'Validate-TestHardPrepared.py'))
baseline,fs=validate['TestValidate'](root,'EnvelopeEnergy.Accepted');assert len(fs)==25
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Native/ControlledPrefix.Result.bin').read_bytes())
assert r.raw[:16]==b'NDMSECPREFIX0001'
context,info,effort,accepted,projection=[r.VectorTest() for _ in range(5)]
assert len(context)==6 and len(info)==6 and info[2]==info[3] and effort[0]<=context[3]
np.testing.assert_array_equal(accepted,source['fields'][-1]);reference=float(context[4]);E=float(source['calibration'][0])
assert context[1:4].tolist()==[source['options'][3],source['options'][4],source['options'][2]]
assert reference==source['lambdas'][-1]
returned=None
if info[0]:returned=[r.VectorTest() for _ in range(4)]
attempts=[[r.VectorTest() for _ in range(4)] for _ in range(int(info[4]))]
pairs=[[r.VectorTest() for _ in range(5)] for _ in range(int(info[5]))]
assert r.position==len(r.raw)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])

def TestNorm(delta):
    """Восстанавливает физическую тензорную меру по129 точкам, без четырёх служебных множителей."""
    strain=np.einsum('spcd,d->spc',op,delta)
    return float(np.sqrt(np.einsum('spc,spc,p,c->',strain,strain,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]))))

def TestWitness(name,field,lam,target):
    """Повторно собирает56 уравнений,4 связи,6 усилий, исходную управляющую строку и обе семьи стали."""
    proof,fields=validate['TestValidate'](root,'ControlledPrefix.'+name+'.Sequence')
    assert len(fields)==26
    np.testing.assert_allclose(fields[:25],source['fields'],rtol=1e-14,atol=1e-300)
    np.testing.assert_allclose(fields[-1],field,rtol=1e-14,atol=1e-300)
    result=proof['states'][-1];assert abs(result['lambda']-lam)<1e-12
    bound=float(context[1]+context[2]*max(abs((L@(source['offset']+lam*source['direction']))[:56])))
    scalar=float(E*(projection@field-target))
    assert max(result['residual_MPa'],result['force_residual_MPa'],abs(scalar))<=bound+2e-11
    result.update(original_scalar_residual_MPa=scalar,original_requested_bound_MPa=bound)
    return result

origin=float(projection@accepted);previous=float(projection@source['fields'][-2]);prior_steps=0;result_attempts=[]
for index,(meta,field,actual,cost) in enumerate(attempts,1):
    assert len(meta)==9 and meta[0]==index and meta[1]==context[5]*.5**(index-1)
    target=origin+meta[1]*(context[0]-origin);ratio=(target-origin)/(origin-previous)
    assert abs(target-meta[2])<1e-18 and abs(ratio-meta[3])<1e-12
    assert 1<=meta[4]<=args.root_budget and cost[0]<=meta[4] and prior_steps+cost[0]<=meta[8]<=context[3]
    item=dict(index=index,fraction=float(meta[1]),target=float(meta[2]),secant_ratio=float(meta[3]),full_root=bool(meta[5]),
        confirmed=bool(meta[6]),total_steps=int(meta[8]),root_effort=cost.tolist())
    if meta[5]:item['root']=TestWitness('Root'+str(index),field,float(meta[7]),target)
    else:assert len(field)==len(actual)==1
    result_attempts.append(item);prior_steps=meta[8]
assert not attempts or prior_steps==effort[0]
result_pairs=[]
for index,(a,full,half,p,meta) in enumerate(pairs,1):
    assert len(meta)==13;np.testing.assert_array_equal(a,accepted);np.testing.assert_array_equal(p,projection)
    full_result=TestWitness('PairFull'+str(index),full,float(meta[0]),float(meta[3]))
    half_result=TestWitness('PairHalf'+str(index),half,float(meta[1]),float(meta[4]))
    length=TestNorm(full-accepted);error=TestNorm(full-(2*half-accepted))/length
    lambda_error=abs(float(2*meta[1]-reference-meta[0]));alpha=meta[10];radius=float(np.hypot(length,alpha*(meta[0]-reference)))
    allowance=(.005*radius)**2-(length*error)**2;noise=8*context[2]*(1+abs(reference))
    bound=max(noise,float(np.sqrt(allowance)/alpha) if allowance>0 else noise)
    joint=float(np.hypot(length*error,alpha*lambda_error)/radius)
    np.testing.assert_allclose([length,error,lambda_error,bound,joint],[meta[9],meta[6],meta[7],meta[8],meta[11]],rtol=1e-8,atol=2e-12)
    assert abs(meta[4]-.5*(origin+meta[3]))<1e-18
    result_pairs.append(dict(full=full_result,half=half_result,field_error=error,lambda_error=lambda_error,lambda_bound=float(bound),
        joint_error=joint,passes_original_half_check=bool(error<=.005 and lambda_error<=bound)))
if returned is not None:
    field,actual,p,uncertainty=returned;assert attempts[-1][0][6] and result_pairs[-1]['passes_original_half_check']
    np.testing.assert_array_equal(p,projection);assert uncertainty[0]>=0 and np.isfinite(uncertainty[0])
    result=TestWitness('Full',field,float(info[1]),float(attempts[-1][0][2]))
    advance=float(projection@(field-accepted));resolution=16*result['original_requested_bound_MPa']/E
    assert abs(advance)>resolution and advance*(attempts[-1][0][2]-origin)>0
else:result=None
proof=dict(native=native,entry_summary_explicitly_checked_separately_from_outer_TOTAL=True,
    actual_components=257,other_parent_components_preserved=257-len(imports),all_working_newton_methods_exact960=True,
    supplied_future_guide=False,source_history25_exact960=True,original25_revision_and_field_unchanged_before_diagnostic_exports=True,
    standalone_diagnostic_budget_not_remaining_full_cold=True,root_step_budget=args.root_budget,shared_step_budget=int(context[3]),total_steps=int(effort[0]),
    attempts=result_attempts,pairs=result_pairs,returned_verified_candidate=bool(info[0]),returned=result,
    native_lambda_uncertainty=float(returned[3][0]) if returned else None,physical_capacity_certified=False,cold_gain_claimed=False)
(root/'ControlledSecantPrefixProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('SECANT PREFIX VERIFIED',bool(info[0]),'ATTEMPTS',len(attempts),'PAIRS',len(pairs),'TOTAL STEPS',int(effort[0]))
