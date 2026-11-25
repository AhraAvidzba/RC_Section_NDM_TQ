"""ДЛЯ ТЕСТОВ. Независимо собирает равновесие точной холодной истории из тензоров трёх станций без нового поиска."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json,re,numpy as np

here=Path(__file__).resolve().parent
p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True);a=p.parse_args();root=a.root
prep=json.loads((root/'Preparation.json').read_bytes());scope=prep['directed_scope']
assert scope['case']=='VerifiedColdStressOnlyReplay' and scope['no_new_NDS_or_capacity_search']
source=root.parent/scope['source_root'];source_proof=json.loads((root/'SourceColdProof.json').read_bytes())
assert (root/'SourceColdProof.json').read_bytes()==(source/scope['source_proof']).read_bytes()
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and audit['imported_this_pass']==['modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==255
for path in (source/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(source)).read_bytes(),path
name='tests/modTestCoupledElasticState.bas'
old=(source/name).read_text('utf-8-sig');new=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub',old)
begin=method.index('    mPtcSuppressTest = False\n')+len('    mPtcSuppressTest = False\n')
end=method.index('    mPolishExecutedTest = True\n',begin)
assert method[begin:end].splitlines()==scope['only_diagnostic_calls_after_frozen_export_removed']
assert old.replace(method,method[:begin]+method[end:],1)==new
raw=(source/'Native/Stock.RawHistory.bin').read_bytes()
assert raw==(root/'Input/SourceColdRawHistory.bin').read_bytes()
frames,raw_info=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Input/SourceColdRawHistory.bin')
assert raw_info['incomplete_tail_bytes']==0;history=frames[-1];count=len(history['fields'])
assert count==scope['source_states']==source_proof['accepted_states']
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
original=read(source/'Input/Reverse1.bin');current=read(root/'Input/Reverse1.bin')
for key in ('calibration','options','offset','direction'):np.testing.assert_array_equal(current[key],original[key])
for key in ('fields','fractions','lambdas'):np.testing.assert_array_equal(current[key],history[key])
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields,history['fields'],rtol=1e-14,atol=1e-300)
assert len(fields)==count
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
strain=np.einsum('spcd,nd->nspc',op,history['fields']);steel=np.einsum('nspc,pc->nsp',strain[:,:,25:],axis[25:])
long=axis[25:,2]==1;assert set(axis[25:,2])=={0.,1.}
limit=np.where(long[None,None,:],np.where(steel>=0,.025,.018),.02)
utilization=abs(steel)/limit;assert np.max(utilization)<=1
energy=np.zeros((count,3,25,2));seen=set()
for row in validate['TestRows'](root/'Native/EnvelopeEnergy.Accepted.csv'):
    if row[0]=='S' and int(row[3])<=25:
        key=(int(row[1])-1,int(row[2])+1,int(row[3])-1);assert key not in seen;seen.add(key)
        energy[key]=np.array([row[10],row[11]],float)
assert len(seen)==count*3*25
cal=np.asarray(current['calibration']);np.testing.assert_allclose(cal,np.array([30000,.2,30,3,16,.002,.12,80,125]),rtol=1e-14,atol=0)
H=np.einsum('nsp,p->n',(cal[6]*energy[:,:,:,0]+cal[7]*energy[:,:,:,1])/cal[8],coords[:25,2])/(3*sum(coords[:25,2]))
for item,lam,fraction in zip(mechanics['states'],history['lambdas'],history['fractions']):
    assert abs(item['lambda']-lam)<=5e-14 and item['fraction']==fraction
    assert max(item['residual_MPa'],item['force_residual_MPa'])<=item['bound_MPa']+1e-12
proof=dict(native=native,actual_components=256,only_test_module_imported=True,all_working_sources_exact_source=True,
    source_root=source.name,all_raw_fields_fractions_and_lambdas_exact=True,source_states=count,
    all_six_forces_and_all_56_physical_and4_gauge_equations=mechanics,
    all_three_station_steel_limits_checked=True,maximum_three_station_steel_utilization=float(np.max(utilization)),
    H_MPa=H.tolist(),H_is_continuation_coordinate_not_failure_criterion=True,
    no_new_NDS_or_capacity_search=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'VerifiedColdStressReplayProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('EXACT COLD STRESS REPLAY VERIFIED',count,'SIX FORCES,56+4 EQUATIONS,THREE-STATION STEEL;NO LIMIT CERTIFICATE')
print('MAX RESIDUAL',max(max(s['residual_MPa'],s['force_residual_MPa']) for s in mechanics['states']),'H LAST',H[-1])
