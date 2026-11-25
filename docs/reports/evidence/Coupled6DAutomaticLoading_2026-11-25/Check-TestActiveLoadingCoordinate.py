"""ДЛЯ ТЕСТОВ. Независимо проверяет исходящую пару CDP и её геометрическую область; не считает спад нагрузки доказательством разрушения."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True)
p.add_argument('--original-predictor-neighborhood',action='store_true');a=p.parse_args();root=a.root
source_root=base/'Cold808PredictorEventCold_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Input/SourceColdRawHistory.bin')
assert raw['incomplete_tail_bytes']==0
for k in ('fields','fractions','lambdas'):assert np.array_equal(source[k],frames[-1][k])
for name in ('src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledDeformationContinuation.cls',
 'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls'):
 assert (root/name).read_bytes()==(source_root/name).read_bytes(),name
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
anchor="' -------------------- ДЛЯ ТЕСТОВ --------------------\n' Использует существующую материальную координату"
assert code.count(anchor)==1 and code.split(anchor)[0].rstrip()==(source_root/name).read_text('utf-8-sig').rstrip()
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=Reader((root/'Native/ActiveLoading.Result.bin').read_bytes());assert r.raw[:16]==b'NDMACTLOAD000001\0'[:16]
meta,accepted,half_seed,full_seed,parameters,half,full,he,fe=[r.VectorTest() for _ in range(9)]
assert r.position==len(r.raw) and len(meta)==14 and len(parameters)==10
assert meta[12]==meta[13] and he[6]==fe[6]==meta[12] and he[0]<=15 and he[0]+fe[0]<=30
assert he[5]==fe[5]==0
assert np.array_equal(accepted,source['fields'][-1]) and meta[2]==source['lambdas'][-1] and meta[3]==source['lambdas'][-2]
assert meta[0]==25 and meta[1]==1
np.testing.assert_allclose(meta[6:8],[meta[4]+.25*(meta[5]-meta[4]),meta[4]+.5*(meta[5]-meta[4])],rtol=1e-15)
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
g,L,w,v,coords,axis,b,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([b+s*.001*gradient for s in (-1,0,1)])
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1.,1.,1.,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0
E,nu=source['calibration'][:2];off=E*nu/((1+nu)*(1-2*nu));G=E/(2*(1+nu))
D=np.diag([off+2*G]*3+[G]*3);D[:3,:3]+=off*(np.ones((3,3))-np.eye(3))
alpha,gamma=parameters[:2];plastic=parameters[2:8];cc,ct=parameters[8:]

def TestLength(delta):
 """Измеряет тот же физический тензорный шаг трёх станций без множителей связей."""
 return float(np.sqrt(max(0.,delta@M@delta)))

def TestYield(field):
 """Считает пробную поверхность25/+1 от захваченной принятой пластической истории."""
 xx,yy,zz,xy,xz,yz=D@((b[24]+.001*gradient[24])@field-plastic)
 tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);tr=np.trace(tensor)
 q=np.sqrt(1.5*np.sum((tensor-np.eye(3)*tr/3)**2));high=np.linalg.eigvalsh(tensor)[-1]
 term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
 return float((alpha*tr+q+term)/(1-alpha)-cc)

np.testing.assert_allclose([TestYield(accepted),TestYield(full_seed)],meta[4:6],rtol=1e-6,atol=2e-12)
rows=validate['TestRows'](root/'Native/ActiveLoading.Yields.csv');assert len(rows)==129
selected=[x for x in rows if x[:3]==['Y','25','1']];assert len(selected)==1
np.testing.assert_allclose(np.array(selected[0][4:7],float),[TestYield(x) for x in (accepted,half_seed,full_seed)],rtol=1e-6,atol=2e-12)
original_distance=TestLength(full_seed-accepted);witnesses={};distances={}
for label,flag,field,factor,target in [('Half',meta[8],half,meta[10],meta[6]),('Full',meta[9],full,meta[11],meta[7])]:
 if len(field)!=60:continue
 distances[label]=TestLength(field-accepted)
 if flag:
  proof,fs=validate['TestValidate'](root,'ActiveLoading.'+label+'.Sequence')
  assert len(fs)==len(source['fields'])+1
  np.testing.assert_allclose(fs[:-1],source['fields'],rtol=1e-14,atol=1e-300)
  np.testing.assert_allclose(fs[-1],field,rtol=1e-14,atol=1e-300)
  bound=.0001*(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+factor*source['direction']))[:56])))
  last=proof['states'][-1]
  assert last['residual_MPa']<=bound+1e-12 and last['force_residual_MPa']<=bound+1e-12
  assert abs(TestYield(field)-target)<=2e-12
  witnesses[label]=dict(equations=last,inner_bound_MPa=float(bound),actual_trial_yield_MPa=TestYield(field),target_MPa=float(target))
pair=None
if len(half)==len(full)==60:
 length=TestLength(full-accepted);error=TestLength(full-(2*half-accepted))/length
 le=abs(2*meta[10]-meta[2]-meta[11]);lb=max(.005*abs(meta[11]-meta[2]),8*.0001*source['options'][4]*(1+abs(meta[2])))
 delta=accepted-source['fields'][-2];outgoing=full-accepted
 slopes=[TestLength(source['fields'][i]-source['fields'][i-1])/abs(source['lambdas'][i]-source['lambdas'][i-1]) for i in range(len(source['fields'])-3,len(source['fields']))]
 metric=float(np.median(slopes));dl=meta[2]-meta[3];out_dl=meta[11]-meta[2]
 mechanical=float(delta@M@outgoing/(TestLength(delta)*length))
 joint=float((delta@M@outgoing+metric*metric*dl*out_dl)/np.sqrt((TestLength(delta)**2+metric**2*dl*dl)*(length**2+metric**2*out_dl*out_dl)))
 limits=dict(Half=(1.5 if a.original_predictor_neighborhood else .75)*original_distance,
  Full=(3. if a.original_predictor_neighborhood else 1.5)*original_distance)
 local=all(distances[k]<=limits[k] for k in limits)
 pair=dict(field_half_error=error,lambda_half_error=float(le),lambda_half_bound=float(lb),
  distances=distances,reference_limits=limits,local=local,
  mechanical_direction_cosine=mechanical,incoming_joint_direction_cosine=joint,
  incoming_load_metric=metric,outgoing_field_per_lambda=length/abs(out_dl),
  both_lambdas_below_source=bool(max(meta[10:12])<meta[2]),
  verified=bool(meta[8] and meta[9] and error<=.005 and le<=lb and local and mechanical>0 and meta[6]>meta[4] and meta[7]>meta[6]))
 if a.original_predictor_neighborhood:assert pair['verified']
 else:assert not meta[8] and not meta[9] and not local
proof=dict(native=native,all_source43_fields_exact=True,source_equilibria_independently_verified=mechanics,
 all_working_methods_unchanged=True,only_test_coordinate_entry_added=True,
 source_point_yield_independently_verified=True,all129station_surface_values_captured=True,
 original_predictor_neighborhood=a.original_predictor_neighborhood,meta=meta.tolist(),
 half_effort=he.tolist(),full_effort=fe.tolist(),shared_primary30_preserved=True,
 accepted_revision_unchanged=True,root_witnesses=witnesses,pair=pair,
 load_decline_does_not_by_itself_certify_physical_capacity=True,
 no_new_continuation_point_accepted=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'ActiveLoadingCoordinateProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root.name,'INDEPENDENT ROOTS',list(witnesses),'PAIR',pair,'PRIMARY',he[0]+fe[0],'NO PHYSICAL LIMIT')
