"""ДЛЯ ТЕСТОВ. Проверяет контексты рабочих дуговых отказов и независимо собирает все фактически найденные корни на точной истории56."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold893Arc56Diagnostics_20261126';parent=base/'Cold888StrictEnvelopeGrowth_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes());assert audit['components']==256 and audit['imported_this_pass']==['modTestCoupledElasticState']
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin');assert len(source['fields'])==56
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Native/Arc56Diagnostics.bin').read_bytes());assert r.raw[:16]==b'NDMARC56DIAG0001'
validate=run_path(str(here/'Validate-TestHardPrepared.py'));initial,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)]);E=source['calibration'][0]
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0
shared=(here/'Check-TestSharedCurveHControl.py').read_text('utf-8-sig');begin=shared.index('def TestLength(');end=shared.index("r=R((root/'Native/HCornerPairs.bin').read_bytes())",begin)
exec(compile(shared[begin:end],str(here/'Check-TestSharedCurveHControl.py'),'exec'),globals())
scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
cases=[]
for index in range(1,9):
    meta=r.VectorTest();assert len(meta)==10 and meta[0]==index and meta[5]==14
    ratio=.5**(index-1) if index<=6 else 1;allowance=12 if index<=6 else (3 if index==7 else 4)
    assert meta[1]==ratio and meta[2]==allowance and meta[4]<=allowance
    packet=[r.VectorTest() for _ in range(14)];info=packet[0];assert len(info)==34
    assert info[14]==info[15] and info[18]==meta[4] and info[16]<=30 and info[17]<=30
    assert bool(meta[3])==bool(info[3]) and info[6]==source['lambdas'][-1] and info[8]==ratio
    np.testing.assert_array_equal(packet[2],source['fields'][-1])
    guides={}
    for name,present in zip(('H','Loading','Tangent'),meta[6:9]):
        header=r.VectorTest();assert len(header)==2 and bool(header[0])==bool(present)
        guides[name]=[r.VectorTest() for _ in range(int(header[1]))]
        assert bool(guides[name])==bool(present)
    effort=r.VectorTest();found=[]
    if info[0]:found.append(('Full',packet[6],info[4],packet[8]))
    if info[1]:found.append(('Half',packet[5],info[5],packet[7]))
    if meta[3]:
        field,actual=r.VectorTest(),r.VectorTest();found.append(('Returned',field,meta[9],actual))
    roots=[]
    for role,field,factor,actual in found:
        phase='Arc56.Case'+str(index)+role+'.Sequence';mechanics,exported=validate['TestValidate'](root,phase)
        assert len(exported)==57
        np.testing.assert_allclose(exported[:56],source['fields'],rtol=1e-14,atol=1e-300);np.testing.assert_allclose(exported[-1],field,rtol=1e-14,atol=1e-300)
        bound=TestBound(factor);weak=mechanics['states'][-1]
        assert weak['residual_MPa']<=bound+1e-12 and weak['force_residual_MPa']<=bound+1e-12
        assert max(abs((actual-source['offset']-factor*source['direction'])/scales))<=bound
        strain=np.einsum('spcd,d->spc',op,field);steel=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
        util=float(max((abs(steel)/np.where(long[None,:],np.where(steel>=0,.025,.018),.02)).ravel()));assert util<=1
        roots.append(dict(role=role,lambda_value=float(factor),weak=weak,H_MPa=TestH(phase,57),maximum_three_station_steel_utilization=util))
    if meta[3]:
        assert info[2] and info[28] in (1,2,3,4)
        if info[28]==1:assert info[0] and info[1] and 0<=info[20]<=.005 and info[21]<=info[22]
    cases.append(dict(case=index,ratio=ratio,call_allowance=allowance,calls=int(meta[4]),solved=bool(meta[3]),
        full_root_found=bool(info[0]),half_root_found=bool(info[1]),pair_verified=bool(info[2]),restored=bool(info[3]),
        pair_iterations=int(info[16]),event_phase_iterations=int(info[17]),total_primary_iterations=int(info[19]),
        event_found=bool(info[26]),selected_kind=int(info[28]),outgoing_tangent_used=bool(info[32]),
        H_guide_called=bool(guides['H']),loading_guide_called=bool(guides['Loading']),tangent_read=bool(guides['Tangent']),
        radius=float(info[11]),fixed_alpha=float(info[9]),field_error=float(info[20]),joint_error=float(info[30]),
        independently_rebuilt_roots=roots,certificate=info.tolist(),guide_headers={k:(v[0].tolist() if v else None) for k,v in guides.items()}))
assert r.position==len(r.raw)
proof=dict(native=native,source_states=56,actual_component_count=256,all_working_sources_exact888=True,cases=cases,
    no_original_history_accepted=True,all_available_roots_independently_equilibrated=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Arc56DiagnosticsProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ARC56 CASES',[(x['ratio'],x['call_allowance'],x['solved'],x['outgoing_tangent_used'],x['H_guide_called'],x['calls']) for x in cases])
