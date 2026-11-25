"""ДЛЯ ТЕСТОВ. Независимо проверяет внутренние полную/половинную H-пробы рабочего алгоритма, включая силы и сталь трёх станций."""
from pathlib import Path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold876VerifiedHCornerPairs20_20261126'
source_file=here/'Check-TestSharedCurveHControl.py';code=source_file.read_text('utf-8-sig');end=code.index("r=R((root/'Native/HCornerPairs.bin').read_bytes())")
code=code[:end]
code=code.replace("root=base/'Cold874SharedCurveHControl20_20261126'","root=base/'Cold876VerifiedHCornerPairs20_20261126'")
code=code.replace("audit['imported_this_pass']==['CCoupledArcContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']","audit['imported_this_pass']==['CCoupledNewtonSolver','modTestCoupledElasticState']")
code=code.replace("len(audit['other_actual_components_exact_parent_except_identifier_case'])==253","len(audit['other_actual_components_exact_parent_except_identifier_case'])==254")
namespace={'__file__':str(source_file),'__name__':'verified_H_pair_source_and_analytic_test'};exec(compile(code,str(source_file),'exec'),namespace)
R=namespace['R'];source=namespace['source'];validate=namespace['validate'];op=namespace['op'];axis=namespace['axis'];coords=namespace['coords'];M=namespace['M'];L=namespace['L'];E=namespace['E']
TestLength=namespace['TestLength'];TestH=namespace['TestH'];TestBound=namespace['TestBound'];accepted=source['fields'][-1];incoming=accepted-source['fields'][-2]
r=R((root/'Native/VerifiedHPairs.bin').read_bytes());assert r.raw[:16]==b'NDMHVERIFIED0001';header=r.VectorTest();assert len(header)==6 and header[0]==20
baseline=TestH('EnvelopeEnergy.Accepted',20);beforeH=TestH('EnvelopeEnergy.Accepted',19);assert abs(header[3]-baseline)<1e-17 and abs(header[4]-beforeH)<1e-17
assert header[1]==source['lambdas'][-1] and header[2]==source['lambdas'][-2] and baseline>beforeH
scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
cases=[];count=0
for index in range(1,4):
    meta=r.VectorTest();effort=r.VectorTest();assert len(meta)==8 and meta[0]==index and meta[1]==.25**(index-1) and meta[6]<=30
    assert abs(meta[2]-(baseline+meta[1]*(baseline-beforeH)))<1e-17
    solved=bool(meta[3]);item=dict(case=index,ratio=float(meta[1]),solved=solved,probes=int(meta[5]),primary_iterations=int(meta[6]))
    if not solved:
        assert meta[7]==0;cases.append(item);continue
    assert meta[7]==1 and meta[5]==2
    witness=[r.VectorTest() for _ in range(6)];info,stored,full,half,full_actual,half_actual=witness;returned=r.VectorTest();actual=r.VectorTest()
    assert len(info)==17 and info[0]==header[1] and info[1]==meta[4] and info[3]==header[3] and info[4]==meta[2] and info[5]==.5*(info[3]+info[4])
    assert info[6]+info[7]==info[8]==meta[6]==effort[0] and info[8]<=30 and info[16]==meta[5]
    np.testing.assert_array_equal(stored,accepted);np.testing.assert_array_equal(returned,full);np.testing.assert_array_equal(actual,full_actual)
    assert returned.tobytes()==full.tobytes() and actual.tobytes()==full_actual.tobytes()
    roots={}
    for role,field,factor,forces,target in (('Full',full,float(info[1]),full_actual,float(info[4])),('Half',half,float(info[2]),half_actual,float(info[5]))):
        assert 0<factor<=source['options'][0]
        phase='VerifiedH.Case'+str(index)+role+'.Sequence';mechanics,fields=validate['TestValidate'](root,phase)
        assert len(fields)==21;np.testing.assert_allclose(fields[:20],source['fields'],rtol=1e-14,atol=1e-300);np.testing.assert_allclose(fields[-1],field,rtol=1e-14,atol=1e-300)
        weak=mechanics['states'][-1];bound=TestBound(factor)
        assert weak['residual_MPa']<=bound+1e-12 and weak['force_residual_MPa']<=bound+1e-12
        force_defect=float(max(abs((forces-source['offset']-factor*source['direction'])/scales)));assert force_defect<=bound
        h=TestH(phase,21);assert h>baseline and abs(h-target)*E/source['calibration'][3]<=max(bound,TestBound(header[1]))+1e-12
        strain=np.einsum('spcd,d->spc',op,field);steel=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
        limits=np.where(long[None,:],np.where(steel>=0,.025,.018),.02);util=float(np.max(abs(steel)/limits));assert util<=1
        length=TestLength(field-accepted);assert length>0;cosine=float(incoming@M@(field-accepted)/(TestLength(incoming)*length))
        roots[role]=dict(lambda_value=factor,H_MPa=h,H_target=target,H_increment=h-baseline,weak=weak,raw_force_residual_MPa=force_defect,
            maximum_three_station_steel_utilization=util,length_over_incoming=length/TestLength(incoming),direction_cosine=cosine)
    length=TestLength(full-accepted);defect=TestLength(full-(2*half-accepted));field_error=defect/length
    le=abs(2*info[2]-info[0]-info[1]);alpha=info[13];distance=np.sqrt(length**2+alpha**2*(info[1]-info[0])**2)
    joint=np.sqrt(defect**2+(alpha*le)**2)/distance;noise=8*.0001*source['options'][4]*(1+abs(info[0]));bound=max(noise,np.sqrt(max(0.,(.005*distance)**2-defect**2))/alpha)
    np.testing.assert_allclose(info[[9,10,11,12]],[field_error,le,bound,joint],rtol=2e-6,atol=1e-11)
    assert field_error<=.005 and le<=bound
    assert bool(info[14])==(roots['Full']['direction_cosine']<=0)
    item.update(roots=roots,field_error=field_error,joint_error=float(joint),lambda_error=float(le),lambda_bound=float(bound),
        full_steps=int(info[6]),half_steps=int(info[7]),actual_turn=bool(info[14]),full_output_exact_after_half=True)
    cases.append(item);count+=1
assert r.position==len(r.raw)
proof=dict(native=namespace['native'],actual_components=256,all14_native_analytical_outcomes_verified=True,true_source20_history_unchanged=True,
    cases=cases,verified_working_H_pairs=count,independently_rebuilt_root_count=2*count,
    absent_failed_roots_not_counted=True,existing_force_tolerances_and_all_three_station_steel_checked=True,
    only_current_benchmark_three_station_limits_proven_not_generic_steel_limit_contract=True,
    generic_runtime_principal_and_deformation_helpers_central_steel_only_audit_open=True,
    no_history_or_physical_capacity_acceptance=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'VerifiedHCornerPairsProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('WORKING VERIFIED H PAIRS',count,'INDEPENDENT FULL/HALF ROOTS',2*count,'CASES',[(x['case'],x['solved'],x['primary_iterations']) for x in cases],';NO PHYSICAL LIMIT')
