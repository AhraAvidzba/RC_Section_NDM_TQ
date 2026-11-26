"""ДЛЯ ТЕСТОВ. Проверяет отдельные материальные истории полного H-шага и двух последовательно принятых половин, не подменяя отсутствие корня разрушением."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold891SequentialAcceptedH_20261126';parent=base/'Cold888StrictEnvelopeGrowth_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes());assert audit['components']==256 and audit['imported_this_pass']==['modTestCoupledElasticState']
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
assert len(source['fields'])==56
validate=run_path(str(here/'Validate-TestHardPrepared.py'));initial,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)]);E=source['calibration'][0]
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0
shared=(here/'Check-TestSharedCurveHControl.py').read_text('utf-8-sig');begin=shared.index('def TestLength(');end=shared.index("r=R((root/'Native/HCornerPairs.bin').read_bytes())",begin)
exec(compile(shared[begin:end],str(here/'Check-TestSharedCurveHControl.py'),'exec'),globals())
r=R((root/'Native/SequentialAcceptedH.bin').read_bytes());assert r.raw[:16]==b'NDMSEQUENTIALH01';header=r.VectorTest();assert len(header)==6 and header[0]==56
baseline=TestH('EnvelopeEnergy.Accepted',56);before=TestH('EnvelopeEnergy.Accepted',55)
assert abs(header[3]-baseline)<1e-17 and abs(header[4]-before)<1e-17 and abs(header[5]-(baseline-before))<1e-17
scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
cases=[];saved={};roots={}
for index in range(1,5):
    meta,effort=r.VectorTest(),r.VectorTest();assert len(meta)==12 and meta[0]==index
    target=baseline+({1:8,2:4,3:8,4:16}[index])*(baseline-before);assert abs(target-meta[2])<1e-17
    solved=bool(meta[3]);attempted=bool(meta[11]);item=dict(case=index,attempted=attempted,solved=solved,primary_iterations=int(meta[6]),probes=int(meta[5]),target_H_MPa=float(meta[2]))
    if not attempted:
        assert not solved and meta[7]==0;cases.append(item);continue
    assert meta[6]<=30
    prefix=source['fields'].copy();factors=source['lambdas'].copy()
    if index in (3,4):
        previous_case={3:2,4:1}[index];assert previous_case in saved
        prefix=np.vstack([prefix,saved[previous_case]['field']]);factors=np.r_[factors,saved[previous_case]['lambda_value']]
    assert len(prefix)==meta[1] and factors[-1]==meta[8]
    if not solved:
        assert meta[7]==0;cases.append(item);continue
    assert meta[7]==1 and meta[5]==2
    info,accepted,full,half,full_actual,half_actual=[r.VectorTest() for _ in range(6)];returned,actual=r.VectorTest(),r.VectorTest()
    assert len(info)==17 and info[0]==meta[8] and info[1]==meta[4] and info[4]==meta[2] and info[5]==.5*(info[3]+info[4])
    assert info[6]+info[7]==info[8]==meta[6]==effort[0]<=30
    np.testing.assert_array_equal(accepted,prefix[-1]);np.testing.assert_array_equal(returned,full);np.testing.assert_array_equal(actual,full_actual)
    end_states={}
    for role,field,factor,forces,Htarget in [('Full',full,info[1],full_actual,info[4]),('Half',half,info[2],half_actual,info[5])]:
        phase='SequentialH.Case'+str(index)+role+'.Sequence';mechanics,exported=validate['TestValidate'](root,phase)
        assert len(exported)==len(prefix)+1
        np.testing.assert_allclose(exported[:-1],prefix,rtol=1e-14,atol=1e-300);np.testing.assert_allclose(exported[-1],field,rtol=1e-14,atol=1e-300)
        weak=mechanics['states'][-1];bound=TestBound(factor)
        assert weak['residual_MPa']<=bound+1e-12 and weak['force_residual_MPa']<=bound+1e-12
        defect=float(max(abs((forces-source['offset']-factor*source['direction'])/scales)));assert defect<=bound
        H=TestH(phase,len(prefix)+1);assert abs(H-Htarget)*E/source['calibration'][3]<=max(bound,TestBound(meta[8]))+1e-12
        tensor=np.einsum('spcd,d->spc',op,field);steel=np.einsum('spc,pc->sp',tensor[:,25:],axis[25:]);long=axis[25:,2]==1
        util=float(max((abs(steel)/np.where(long[None,:],np.where(steel>=0,.025,.018),.02)).ravel()));assert util<=1
        end_states[role]=dict(lambda_value=float(factor),H_MPa=H,weak=weak,raw_force_residual_MPa=defect,maximum_three_station_steel_utilization=util)
    length=TestLength(full-accepted);curve_defect=TestLength(full-2*half+accepted);field_error=curve_defect/length
    le=abs(2*info[2]-info[0]-info[1]);alpha=info[13];distance=np.sqrt(length**2+alpha**2*(info[1]-info[0])**2)
    joint=np.sqrt(curve_defect**2+(alpha*le)**2)/distance
    noise=8*.0001*source['options'][4]*(1+abs(info[0]));bound=max(noise,np.sqrt(max(0,(.005*distance)**2-curve_defect**2))/alpha)
    assert field_error<=.005 and le<=bound
    np.testing.assert_allclose(info[[9,10,11,12]],[field_error,le,bound,joint],rtol=2e-6,atol=1e-11)
    item.update(roots=end_states,field_error=field_error,joint_error=float(joint),full_steps=int(info[6]),half_steps=int(info[7]))
    saved[index]=dict(field=full.copy(),lambda_value=float(info[1]),phase='SequentialH.Case'+str(index)+'Full.Sequence',count=len(prefix)+1)
    cases.append(item)
assert r.position==len(r.raw)
comparison=None
if 1 in saved and 3 in saved:
    a,b=saved[1],saved[3];delta=a['field']-source['fields'][-1];difference=b['field']-a['field']
    def TestLastMaterial(case):
        """Читает фактические тензоры, энергетические координаты и сталь только конечной принятой истории пути."""
        values={}
        for row in validate['TestRows'](root/'Native'/(case['phase']+'.csv')):
            if row[0]=='S' and int(row[1])==case['count']:values[(int(row[2]),int(row[3]))]=np.array(row[4:],float)
        assert len(values)==129;return values
    aa,bb=TestLastMaterial(a),TestLastMaterial(b);diff=np.array([bb[k]-aa[k] for k in sorted(aa)])
    comparison=dict(one_full_step_H_ratio=8,two_sequential_accepted_half_H_ratios=[4,4],
        relative_field_difference=TestLength(difference)/TestLength(delta),lambda_difference=b['lambda_value']-a['lambda_value'],
        maximum_concrete_stress_difference_MPa=float(np.max(abs(diff[:,:6]))),maximum_bar_stress_difference_MPa=float(np.max(abs(diff[:,-1]))),
        maximum_energy_coordinate_difference=float(np.max(abs(diff[:,6:8]))),
        full_internal_material_state_vector_not_compared=True,not_an_automatic_acceptance_or_failure_criterion=True)
proof=dict(native=native,source_states=56,cases=cases,verified_pairs=sum(x['solved'] for x in cases),
    sequential_history_comparison=comparison,all_working_sources_exact888=True,independent_weak_and_force_assembly=True,
    all_three_station_steel_checked=True,intermediate_fields_replayed_as_accepted=True,original56_restored_at_end=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'SequentialAcceptedHProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('SEQUENTIAL ACCEPTED H CASES',[(x['case'],x['solved'],x['primary_iterations']) for x in cases],'COMPARISON',comparison)
