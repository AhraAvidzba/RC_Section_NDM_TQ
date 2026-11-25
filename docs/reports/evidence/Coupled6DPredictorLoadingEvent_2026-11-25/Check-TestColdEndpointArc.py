"""ДЛЯ ТЕСТОВ. Проверяет точный холодный конец, независимые уравнения и сертификат одной дуговой операции без требования её успеха."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json,struct
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True)
p.add_argument('--source',type=Path,default=base/'Cold802NestedMaterialAccuracy_20261125')
p.add_argument('--tighter-state',action='store_true');p.add_argument('--predictor-event',action='store_true');a=p.parse_args();root=a.root
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
allowance_count,allowance=struct.unpack('<id',(root/'Input/ArcCallAllowance.bin').read_bytes())
assert allowance_count==1 and allowance>=3 and allowance==int(allowance)
history,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](a.source/'Native/Stock.RawHistory.bin')
assert raw['incomplete_tail_bytes']==0
for k in ('fields','fractions','lambdas'):assert np.array_equal(source[k],history[-1][k]),k
for name in ('src/Solver/CCoupledNewtonSolver.cls',
 'src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteReturn.cls'):
 assert (root/name).read_bytes()==(a.source/name).read_bytes(),name
name='src/Solver/CCoupledArcContinuation.cls';arc_code=(root/name).read_text('utf-8-sig')
if a.predictor_event:
 import ast
 tree=ast.parse((here/'Stage-TestPredictorLoadingEvent.py').read_text('utf-8-sig'))
 replacements={node.targets[0].id:ast.literal_eval(node.value) for node in tree.body if isinstance(node,ast.Assign)
  and isinstance(node.targets[0],ast.Name) and node.targets[0].id in ('old','new')}
 assert arc_code.count(replacements['new'])==1
 arc_code=arc_code.replace(replacements['new'],replacements['old'],1)
assert arc_code==(a.source/name).read_text('utf-8-sig')
name='src/Materials/CCoupledConcreteResponse.cls';response=(root/name).read_text('utf-8-sig')
if a.tighter_state:
 new='Private Const STATE_TOLERANCE As Double = 0.0000000000005'
 old='Private Const STATE_TOLERANCE As Double = 0.0000000005'
 assert response.count(new)==1;response=response.replace(new,old,1)
assert response==(a.source/name).read_text('utf-8-sig')
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=Reader((root/'Native/OperationalArc.Result.bin').read_bytes());assert r.raw[:16]==b'NDMOPERARC000003'
packet=[r.VectorTest() for _ in range(15)]
info,projection,accepted,half_seed,full_seed,half,full,ah,af,event,face,face_half,unused,ah_face,parameters=packet
returned=actual=returned_projection=None
if info[3]:returned,actual,returned_projection=[r.VectorTest() for _ in range(3)]
effort=r.VectorTest();assert r.position==len(r.raw)
assert len(info)==30 and info[14]==info[15] and info[16]<=30 and info[17]<=30 and info[18]<=allowance
assert info[19]==(info[16] if info[28]==1 else info[16]+info[17])
assert np.array_equal(accepted,source['fields'][-1]) and info[6]==source['lambdas'][-1]
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
n=len(fields);E=source['calibration'][0]
ops=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
g,L,w,v,coords,axis,b,gradient=ops
op=np.array([b+s*.001*gradient for s in (-1,0,1)])
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1.,1.,1.,.5,.5,.5]))/(3*sum(coords[:,2]))
M[56:]=0;M[:,56:]=0

def TestLength(delta):
 """Измеряет физическое поле в исходной трёхстанционной метрике, исключая четыре связи."""
 return float(np.sqrt(max(0.,delta@M@delta)))

def TestBound(factor):
 """Восстанавливает фактический внутренний допуск без округления усилий LoadPath."""
 return float(.0001*(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+factor*source['direction']))[:56]))))

def TestWitness(label,field,factor):
 """Собирает все уравнения по сохранённым напряжениям; отдельный материал заново не изобретается."""
 proof,fs=validate['TestValidate'](root,'OperationalArc.'+label+'.Sequence')
 assert len(fs)==n+1
 np.testing.assert_allclose(fs[:n],source['fields'],rtol=1e-14,atol=1e-300)
 np.testing.assert_allclose(fs[n],field,rtol=1e-14,atol=1e-300)
 assert abs(proof['states'][-1]['lambda']-factor)<1e-12
 assert proof['states'][-1]['residual_MPa']<=TestBound(factor)+1e-12
 return proof['states'][-1]

event_check=None
if info[26]:
 assert len(event)==7 and len(parameters)==10 and event[6]==info[14]
 point,side=int(event[1]),int(event[2]);alpha,gamma=parameters[:2];plastic=parameters[2:8];cc,ct=parameters[8:]
 nu=source['calibration'][1];off=E*nu/((1+nu)*(1-2*nu));G=E/(2*(1+nu))
 D=np.diag([off+2*G]*3+[G]*3);D[:3,:3]+=off*(np.ones((3,3))-np.eye(3))
 def TestYield(field):
  """Проверяет выбранную поверхность по исходной пластической истории, независимо от признаков ветви VBA."""
  xx,yy,zz,xy,xz,yz=D@((b[point-1]+side*.001*gradient[point-1])@field-plastic)
  tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);tr=np.trace(tensor)
  q=np.sqrt(1.5*np.sum((tensor-np.eye(3)*tr/3)**2));high=np.linalg.eigvalsh(tensor)[-1]
  term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
  return float((alpha*tr+q+term)/(1-alpha)-cc)
 bracket=full if info[0] else half if info[1] else full_seed
 np.testing.assert_allclose([TestYield(accepted),TestYield(bracket)],event[3:5],rtol=1e-5,atol=2e-12)
 assert abs(TestYield(face))<=max(2e-12,TestBound(info[27]))
 event_check=dict(point=point,side=side,face_yield_MPa=TestYield(face),lambda_value=float(info[27]))
 if len(face_half)==60:
  spacing=TestLength(face-accepted);error=TestLength(face-(2*face_half-accepted))/spacing
  le=abs(2*info[24]-info[6]-info[27]);lb=max(.005*abs(info[27]-info[6]),8*.0001*source['options'][4]*(1+abs(info[6])))
  np.testing.assert_allclose(info[25],error,rtol=1e-5,atol=1e-12)
  event_check.update(field_half_error=error,lambda_half_error=le,lambda_half_bound=lb,
   half_yield_MPa=TestYield(face_half),half_resolution_MPa=.5*spacing*E,
   required_half_resolution_MPa=8*max(TestBound(info[27]),TestBound(info[24])))
  if info[28]==2:
   assert error<=.005 and le<=lb and TestYield(face_half)<0
   assert event_check['half_resolution_MPa']>event_check['required_half_resolution_MPa']

witnesses={}
for label,flag,field,factor in (('PairFull',info[0],full,info[4]),('PairHalf',info[1],half,info[5]),
 ('Event',info[26],face,info[27]),('EventHalf',len(face_half)==60,face_half,info[24]),
 ('Full',info[3],returned,info[23])):
 if flag:witnesses[label]=TestWitness(label,field,factor)
pair=None
if info[0] and info[1]:
 length=TestLength(full-accepted)
 error=TestLength(full-(2*half-accepted))/length
 le=abs(2*info[5]-info[6]-info[4]);lb=max(.005*abs(info[4]-info[6]),8*.0001*source['options'][4]*(1+abs(info[6])))
 np.testing.assert_allclose(info[21:23],[le,lb],rtol=1e-5,atol=1e-15)
 if info[28] in (0,1):np.testing.assert_allclose(info[20],error,rtol=1e-5,atol=1e-12)
 resolved=.5*info[11]*E>8*max(TestBound(info[6]),TestBound(info[4]),TestBound(info[5]))
 local=[]
 for field,factor,radius in ((full,info[4],info[11]),(half,info[5],.5*info[11])):
  distance=np.hypot(TestLength(field-accepted),info[9]*(factor-info[6]))
  progress=projection@(field-accepted)+info[12]*(factor-info[6])
  local.append(bool(factor>0 and distance<=3*radius and progress>0))
 pair=dict(field_error=error,lambda_error=le,lambda_bound=lb,resolved=bool(resolved),local=local,
  verified=bool(error<=.005 and le<=lb and resolved and all(local)))
 if info[28]==1:assert pair['verified']
approach_reader=Reader((root/'Native/OperationalArc.Approaches.bin').read_bytes())
assert approach_reader.raw[:16]==b'NDMARCAPPR000001'
count=approach_reader.VectorTest();assert count[0]==info[29] and 0<=count[0]<=4
approaches=[]
for j in range(1,int(count[0])+1):
 record=[approach_reader.VectorTest() for _ in range(6)];meta,candidate,probe,p_event,ac,ap=record
 assert meta[0]==j and meta[11]==meta[12]==info[14] and meta[13]<=30 and meta[15]<allowance
 item=dict(meta=meta.tolist(),candidate=TestWitness('ApproachCandidate'+str(j),candidate,meta[2]))
 if len(probe)==60:
  item['probe']=TestWitness('ApproachProbe'+str(j),probe,meta[3])
  length=TestLength(candidate-accepted);error=TestLength(candidate-(2*probe-accepted))/length
  le=abs(2*meta[3]-info[6]-meta[2]);lb=max(.005*abs(meta[2]-info[6]),8*.0001*source['options'][4]*(1+abs(info[6])))
  np.testing.assert_allclose(meta[5:9],[length,error,le,lb],rtol=1e-5,atol=1e-14)
  if meta[10]:assert error<=.005 and le<=lb and bool(meta[9])
 approaches.append(item)
assert approach_reader.position==len(approach_reader.raw)
if info[3]:
 assert info[2] and effort[0]==0 and info[28] in (1,2,3)
 if info[28]==3:assert approaches and approaches[-1]['meta'][10]
proof=dict(native=native,source_cold_fields_exact=True,source_states=n,source_name=a.source.name,
 material_state_tolerance_tightened_only=a.tighter_state,predictor_event_fallback_only=a.predictor_event,source_all_equations_independent=mechanics,
 accepted_revision_unchanged=True,info=info.tolist(),event=event.tolist(),effort=effort.tolist(),
 arc_call_allowance=int(allowance),
 returned_verified_candidate=bool(info[3]),pair=pair,event_check=event_check,approaches=approaches,root_witnesses=witnesses,
 all_primary_and_call_budgets_preserved=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'ColdEndpointArcProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root.name,'SOURCE',n,'INDEPENDENT EQUILIBRIA; ARC',bool(info[3]),'PAIR',bool(info[0]),bool(info[1]),
 'PRIMARY',info[19],'CALLS',info[18],'LAST RESIDUAL',effort[-1],'NO PHYSICAL LIMIT')
