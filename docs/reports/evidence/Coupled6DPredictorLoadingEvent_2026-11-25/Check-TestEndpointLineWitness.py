"""ДЛЯ ТЕСТОВ. Сравнивает фактические ответы вдоль отвергнутой поправки с захваченной касательной; не принимает неравновесные пробы."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold805EndpointLineWitness_20261125';parent=base/'Cold804HistoryAccuracyEndpoint_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
for name in ('OperationalArc.Result.bin','OperationalArc.Approaches.bin'):
 assert (root/'Native'/name).read_bytes()==(parent/'Native'/name).read_bytes(),name
for name in ('CCoupledConcreteResponse.cls','CCoupledConcreteReturn.cls'):
 assert (root/'src/Materials'/name).read_bytes()==(parent/'src/Materials'/name).read_bytes(),name
read=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=read((root/'Native/EndpointLine.Result.bin').read_bytes());assert r.raw[:16]==b'NDMENDLINE000001\0'[:16]
meta,field,correction,residual,flat,projection,load_direction=[r.VectorTest() for _ in range(7)]
fractions=r.VectorTest();assert r.position==len(r.raw)
assert list(map(len,(meta,field,correction,residual,flat,projection,load_direction)))==[10,60,61,60,3600,60,60]
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
E=source['calibration'][0];matrix=flat.reshape(60,60).T
predicted=matrix@correction[:60]-load_direction*correction[60]
linear_error=float(max(abs(predicted+residual)))
control_error=float(abs(E*(projection[:56]@correction[:56]+meta[2]*correction[60])+meta[9]))
assert linear_error<1e-14 and control_error<1e-14,(linear_error,control_error)
rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows']
g,L,w,v,coords,axis,b,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](
 root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
assert np.array_equal(fractions,[-1,0,1e-6,1e-4,.01,.1,1])
samples=[];networks=[];residuals=[]
n=len(source['fields']);half=.001
for sample,fraction in enumerate(fractions,1):
 records=rows(root/f'Native/EndpointLine.Sample{sample}.Sequence.csv')
 fs=[x for x in records if x[0]=='F'];assert len(fs)==n+1
 for i,x in enumerate(fs[:n]):np.testing.assert_allclose(np.array(x[4:],float),source['fields'][i],rtol=1e-14,atol=1e-300)
 final=fs[-1];factor=float(final[3]);actual_field=np.array(final[4:],float)
 np.testing.assert_allclose(actual_field,field+fraction*correction[:60],rtol=1e-14,atol=1e-300)
 assert abs(factor-(meta[0]+fraction*correction[60]))<1e-14
 stresses=np.zeros((3,43,6));bars=np.zeros((3,43))
 for kind,*x in records:
  if kind=='S' and int(x[0])==n+1:
   side,point=int(x[1])+1,int(x[2])-1;stresses[side,point]=np.array(x[3:9],float);bars[side,point]=float(x[11])
 net=stresses.copy();net[:,25:]=bars[:,25:,None]*axis[None,25:]-stresses[:,25:]
 point=np.einsum('pcd,pc->pd',w,net[1])+np.einsum('pcd,pc->pd',v,(net[2]-net[0])/(2*half))
 R=g@actual_field+point.sum(axis=0)-L@(source['offset']+factor*source['direction'])
 native_row=[x for x in records if x[0]=='R'][-1];norm=max(max(abs(R[:56])),E*max(abs(R[56:])))
 assert abs(norm-float(native_row[2]))<2e-12,(sample,norm,native_row[2])
 residuals.append(R);networks.append(net)
 samples.append(dict(sample=sample,fraction=float(fraction),factor=factor,weak_norm_MPa=float(norm),native_weak_norm_MPa=float(native_row[2])))
base_R=residuals[1]
assert max(abs(base_R-residual))<2e-12
row=int(np.argmax(abs(residual[:56])))
for index,item in enumerate(samples):
 delta=residuals[index]-base_R
 item.update(row=row+1,row_residual_MPa=float(residuals[index][row]),actual_row_change_MPa=float(delta[row]),
  predicted_row_change_MPa=float(item['fraction']*predicted[row]),
  field_linearization_defect_MPa=float(max(abs((delta-item['fraction']*predicted)[:56]))))
delta_net=networks[-1]-networks[1]
delta_point=np.einsum('pcd,pc->pd',w,delta_net[1])+np.einsum('pcd,pc->pd',v,(delta_net[2]-delta_net[0])/(2*half))
top=sorted((dict(point=i+1,row_change_MPa=float(delta_point[i,row]),
 largest_stress_change_MPa=float(np.max(abs(delta_net[:,i])))) for i in range(43)),key=lambda x:abs(x['row_change_MPa']),reverse=True)[:10]
proof=dict(native=native,original804_arc_packet_bitidentical=True,material_classes_exact804=True,
 captured_meta=meta.tolist(),linear_solve_error_MPa=linear_error,control_linear_error_MPa=control_error,
 capture_residual_independently_rebuilt=True,all_source45_fields_exact=True,
 samples=samples,governing_row=row+1,predicted_row_slope_MPa=float(predicted[row]),top_point_contributions=top,
 samples_are_post_attempt_material_replays_not_accepted_continuation=True,
 physical_capacity_certified=False,whole_goal_completed=False)
(root/'EndpointLineWitnessProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('805 ORIGINAL804 ARC BITIDENTICAL; LINEAR ERROR',linear_error,'CONTROL',control_error,'ROW',row+1)
for sample in samples:print('SAMPLE',sample)
print('POINT CONTRIBUTIONS',top[:4])
