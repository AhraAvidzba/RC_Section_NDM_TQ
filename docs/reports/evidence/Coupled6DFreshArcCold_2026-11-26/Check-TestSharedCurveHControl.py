"""ДЛЯ ТЕСТОВ. Независимо проверяет общую метрику на аналитических максимумах и все фактические H-корни без принятия отсутствующих половин."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold874SharedCurveHControl20_20261126';parent=base/'Cold867FixedMetricKernel_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay');assert native==dict(passed=3,failed=0,macros=3,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes());assert audit['components']==256 and audit['imported_this_pass']==['CCoupledArcContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==253
for f in (parent/'src').rglob('*'):
    if f.is_file() and f.name not in ('CCoupledArcContinuation.cls','CCoupledNewtonSolver.cls'):assert f.read_bytes()==(root/f.relative_to(parent)).read_bytes()
arc=(root/'src/Solver/CCoupledArcContinuation.cls').read_text('utf-8-sig');assert arc.count('PairLambdaBound = mSolver.CurveLambdaBound(')==1
newton=(root/'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig');assert newton.count('lambdaBound = CurveLambdaBound(')==2
assert newton.count('Private Function ConfirmOriginalReturnHalf(')==1 and newton.count('Friend Function CurveLambdaBound(')==1
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Native/SharedCurveCases.bin').read_bytes());assert r.raw[:16]==b'NDMSHAREDCURVE01'
count=r.VectorTest();assert count[0]==14;analytic=[]
for index in range(14):
    x=r.VectorTest();out=r.VectorTest();length,error,factor,reference,le,alpha,limit,relative=x
    distance=np.sqrt(length**2+alpha**2*(factor-reference)**2);defect=error*length;noise=8*relative*(1+abs(reference))
    bound=max(noise,np.sqrt(max(0.,(limit*distance)**2-defect**2))/alpha)
    joint=np.sqrt(defect**2+(alpha*le)**2)/distance
    step=max(error,np.sqrt(defect**2+(alpha*(0 if le<=noise else le))**2)/distance)
    np.testing.assert_allclose(out,[bound,joint,step],rtol=2e-14,atol=1e-16)
    admit=error<=limit and le<=bound;old=error<=limit and le<=max(noise,limit*abs(factor-reference))
    assert admit==(index!=13)
    if index<12:assert not old
    analytic.append(dict(case=index+1,new_metric_admits=bool(admit),old_metric_admits=bool(old),joint_error=float(joint),lambda_error=float(le),lambda_bound=float(bound)))
assert r.position==len(r.raw)
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];source=read(root/'Input/Reverse1.bin');assert len(source['fields'])==20
assert (root/'Input/Reverse1.bin').read_bytes()==(base/'Cold869FixedMetricNextScan_20261126/Input/Reverse1.bin').read_bytes()
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)]);E=source['calibration'][0]
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0

def TestLength(delta):
    """Использует тензорную норму трёх станций, не включая реакции связей."""
    return float(np.sqrt(max(0.,delta@M@delta)))

def TestH(phase,count):
    """Считает принятую численную материальную координату, не подменяя её повреждением или критерием разрушения."""
    gt,gc,ell=source['calibration'][6:9];total=0.;seen=set()
    for row in validate['TestRows'](root/'Native'/(phase+'.csv')):
        if row[0]=='S' and int(row[1])==count and int(row[3])<=25:
            key=int(row[2]),int(row[3]);assert key not in seen;seen.add(key)
            total+=coords[key[1]-1,2]*(gt*float(row[10])+gc*float(row[11]))/ell
    assert len(seen)==75
    return total/(3*sum(coords[:25,2]))

def TestBound(factor):
    """Возвращает исходный строгий силовой допуск этих H-проб в фактических единицах МПа."""
    return float(.0001*(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+factor*source['direction']))[:56]))))

r=R((root/'Native/HCornerPairs.bin').read_bytes());assert r.raw[:16]==b'NDMHCORNERPAIR01';header=r.VectorTest();assert len(header)==8
assert header[0]==20 and header[1]==source['lambdas'][-1] and header[2]==source['lambdas'][-2]
baseline=TestH('EnvelopeEnergy.Accepted',20);beforeH=TestH('EnvelopeEnergy.Accepted',19)
assert abs(header[3]-baseline)<1e-17 and abs(header[4]-beforeH)<1e-17 and baseline>beforeH
alpha=float(header[6]);expected_alpha=json.loads((parent/'FixedMetricKernelProof.json').read_bytes())['independent_elastic_alpha'];assert abs(alpha-expected_alpha)<2e-13*alpha
accepted=source['fields'][-1];incoming=accepted-source['fields'][-2];scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
records=[];pairs=[];solved_count=0
for case in range(1,4):
    used=0;case_roots={}
    for role in (1,2):
        meta=r.VectorTest();effort=r.VectorTest();field=r.VectorTest();actual=r.VectorTest();energy=r.VectorTest()
        assert len(meta)==13 and tuple(meta[:2])==(case,role) and meta[2]==.25**(case-1) and meta[3]==(1. if role==1 else .5)
        assert meta[10]==meta[11]==header[7] and meta[12]==30-used and meta[8]>=0
        used+=int(meta[8]);assert meta[9]==used<=30
        ready=bool(meta[5]);item=dict(case=case,role='Full' if role==1 else 'Half',attempted=bool(meta[12]>0),solved=ready,iterations=int(meta[8]),total_pair_iterations=used,probes=int(meta[7]))
        if not ready:
            assert len(field)==len(actual)==1 and energy[0]==0
            if meta[12]==0:assert meta[8]==meta[7]==0
            records.append(item);continue
        assert meta[12]>0 and effort[0]==meta[8] and len(field)==60 and len(actual)==6
        factor=float(meta[6]);assert 0<factor<=source['options'][0]
        phase='HCorner.Case'+str(case)+'Role'+str(role)+'.Sequence';mech,ff=validate['TestValidate'](root,phase)
        assert len(ff)==21;np.testing.assert_allclose(ff[:20],source['fields'],rtol=1e-14,atol=1e-300);np.testing.assert_allclose(ff[-1],field,rtol=1e-14,atol=1e-300)
        weak=mech['states'][-1];bound=TestBound(factor)
        assert weak['residual_MPa']<=bound+1e-12 and weak['force_residual_MPa']<=bound+1e-12
        raw_force=float(max(abs((actual-source['offset']-factor*source['direction'])/scales)));assert raw_force<=bound
        h=TestH(phase,21);assert h>baseline and abs(h-energy[0])<1e-17
        assert abs(h-meta[4])*E/source['calibration'][3]<=bound+1e-12
        strains=np.einsum('spcd,d->spc',op,field);steel=np.einsum('spc,pc->sp',strains[:,25:],axis[25:]);long=axis[25:,2]==1
        limits=np.where(long[None,:],np.where(steel>=0,.025,.018),.02);util=float(np.max(abs(steel)/limits));assert util<=1
        length=TestLength(field-accepted);assert length>0
        cosine=float(incoming@M@(field-accepted)/(TestLength(incoming)*length))
        item.update(lambda_value=factor,H_MPa=h,H_increment=h-baseline,H_target=float(meta[4]),weak=weak,raw_force_residual_MPa=raw_force,steel_utilization=util,direction_cosine=cosine)
        case_roots[role]=(field,factor);solved_count+=1;records.append(item)
    pair=dict(case=case,full_solved=1 in case_roots,half_solved=2 in case_roots)
    if len(case_roots)==2:
        full,lf=case_roots[1];half,lh=case_roots[2];reference=float(header[1]);length=TestLength(full-accepted)
        defect=TestLength(full-(2*half-accepted));field_error=defect/length;le=abs(2*lh-reference-lf)
        distance=np.sqrt(length**2+alpha**2*(lf-reference)**2);joint=np.sqrt(defect**2+(alpha*le)**2)/distance
        noise=8*.0001*source['options'][4]*(1+abs(reference));lb=max(noise,np.sqrt(max(0.,(.005*distance)**2-defect**2))/alpha)
        pair.update(field_error=field_error,joint_error=float(joint),lambda_error=le,lambda_bound=float(lb),new_curve_qualifies=bool(field_error<=.005 and le<=lb),
                    old_curve_qualifies=bool(field_error<=.005 and le<=max(noise,.005*abs(lf-reference))))
    pairs.append(pair)
assert r.position==len(r.raw)
proof=dict(native=native,actual_components=256,all14_native_analytical_outcomes_verified=True,analytic=analytic,
    true20_history_unchanged=True,shared_curve_metric_delegation_and_two_Newton_consumers_verified=True,
    full_H_roots=sum(x['solved'] for x in records if x['role']=='Full'),half_H_roots=sum(x['solved'] for x in records if x['role']=='Half'),
    total_actual_solved_roots=solved_count,roots=records,pairs=pairs,original_full30_budget_preserved=True,
    all_existing_roots_independently_rebuilt_count=solved_count,absent_or_unattempted_half_roots_not_counted=True,
    no_history_or_capacity_acceptance=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'SharedCurveHControlProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ANALYTICAL14 VERIFIED;FULL H ROOTS',proof['full_H_roots'],'HALF',proof['half_H_roots'],'PAIRS',pairs,'NO PHYSICAL LIMIT')
