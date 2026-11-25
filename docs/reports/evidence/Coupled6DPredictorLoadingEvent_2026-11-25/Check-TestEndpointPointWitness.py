"""ДЛЯ ТЕСТОВ. Подтверждает пересечение начала нагружения25/+1 и отделяет его от ошибки линейного решения или локального допуска."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold806EndpointPointWitness_20261125';parent=base/'Cold805EndpointLineWitness_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
for name in ('OperationalArc.Result.bin','OperationalArc.Approaches.bin','EndpointLine.Result.bin'):
 assert (root/'Native'/name).read_bytes()==(parent/'Native'/name).read_bytes(),name
for name in ('CCoupledConcreteResponse.cls','CCoupledConcreteReturn.cls'):
 assert (root/'src/Materials'/name).read_bytes()==(parent/'src/Materials'/name).read_bytes()
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=Reader((root/'Native/EndpointPoint.Result.bin').read_bytes());assert r.raw[:16]==b'NDMENDPOINT00001\0'[:16]
points,fs=r.VectorTest(),r.VectorTest();assert np.array_equal(points,[25,21,16,10])
arc=Reader((root/'Native/OperationalArc.Result.bin').read_bytes());packet=[arc.VectorTest() for _ in range(15)]
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
b,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)[-2:]
E,nu=source['calibration'][:2];off=E*nu/((1+nu)*(1-2*nu));G=E/(2*(1+nu))
D=np.diag([off+2*G]*3+[G]*3);D[:3,:3]+=off*(np.ones((3,3))-np.eye(3))
records=[];selected=None
for point in points:
 for side in (-1,0,1):
  meta,state,strain,de=[r.VectorTest() for _ in range(4)]
  assert meta[0]==point and meta[1]==side and meta[2]==packet[0][14]
  modes=[[r.VectorTest() for _ in range(3)] for _ in range(2)]
  trials=[[r.VectorTest() for _ in range(2)] for _ in fs]
  alpha,gamma=state[:2];plastic=state[2:8];cc,ct=state[8:]
  def TestYield(epsilon):
   """Восстанавливает ту же поверхность из захваченной принятой истории и независимо вычисленных инвариантов."""
   xx,yy,zz,xy,xz,yz=D@(epsilon-plastic)
   tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);tr=np.trace(tensor)
   q=np.sqrt(1.5*np.sum((tensor-np.eye(3)*tr/3)**2));high=np.linalg.eigvalsh(tensor)[-1]
   term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
   return float((alpha*tr+q+term)/(1-alpha)-cc)
  for fraction,(context,stress) in zip(fs,trials):
   assert context[0]==fraction and context[1] and context[5] and context[7]==0
   assert abs(TestYield(strain+fraction*de)-context[6])<2e-12
  actual=trials[-1][1]-trials[1][1]
  item=dict(point=int(point),side=side,state=state.tolist(),strain=strain.tolist(),increment=de.tolist(),
   branches=[x[0].tolist() for x in trials],actual_full_stress_change_MPa=actual.tolist(),
   tangent_modes=[dict(mode=int(x[0][0]),code=int(x[0][1]),predicted_stress_change_MPa=(x[2].reshape(6,6)@de).tolist(),
    full_stress_error_MPa=float(max(abs(x[2].reshape(6,6)@de-actual)))) for x in modes])
  assert all(x[0][1]==0 for x in modes)
  if point==25 and side==1:
   assert trials[1][0][2]==0 and trials[-1][0][2]==1
   assert trials[1][0][6]<0<trials[-1][0][6]
   assert all(x[0][3]==1 and x[0][4]==1 for x in trials)
   predicted_yields={name:TestYield((b[24]+.001*gradient[24])@field) for name,field in zip(('accepted','half_seed','full_seed'),packet[2:5])}
   assert predicted_yields['accepted']<0<predicted_yields['half_seed']<predicted_yields['full_seed']
   item['source_and_predictor_yields_MPa']=predicted_yields
   item['crossing_fraction_linear_estimate']=-trials[1][0][6]/(trials[-1][0][6]-trials[1][0][6])
   selected=item
  records.append(item)
assert r.position==len(r.raw) and selected is not None
assert not packet[0][0] and not packet[0][1] and packet[0][18]==2 and not packet[0][26]
proof=dict(native=native,original805_attempt_and_direction_bitidentical=True,
 all_trial_yield_values_independently_verified=True,point25_plus1_loading_transition_confirmed=True,
 both_arc_roots_failed_before_event_attempt=True,predictor_brackets_real_loading_transition=True,
 records=records,selected_point=selected,physical_capacity_certified=False,whole_goal_completed=False)
(root/'EndpointPointWitnessProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('806 REAL LOADING CROSSING25/+1; PREDICTOR BRACKETS, BOTH ARC ROOTS FAIL, EVENT NOT ATTEMPTED')
print('YIELDS',selected['source_and_predictor_yields_MPa'],'CROSSING FRACTION',selected['crossing_fraction_linear_estimate'])
