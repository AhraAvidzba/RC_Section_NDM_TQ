"""ДЛЯ ТЕСТОВ. Проверяет порядок истории и упругую разгрузку в отмеченной петле материальной точки."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json, numpy as np
repo=Path.cwd(); here=Path(__file__).resolve().parent
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold937MaterialDiagramView_20261126')
v=run_path(str(here/'Validate-TestHardPrepared.py'))
rows=v['TestRows'](root/'Native/EnvelopeEnergy.Accepted.csv')
frames={int(r[1]):np.array(r[4:],float) for r in rows if r[0]=='F'}
stresses={int(r[1]):r for r in rows if r[0]=='S' and int(r[2])==-1 and int(r[3])==11}
ops=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
B,gradient=ops[-2:]; operator=(B-.001*gradient)[10]
E=30000.; nu=.2; G=E/(2*(1+nu)); lame=E*nu/((1+nu)*(1-2*nu))
D=np.zeros((6,6)); D[:3,:3]=lame; D[np.arange(3),np.arange(3)]+=2*G; D[3:,3:]=np.eye(3)*G
def TestTensor(s):
    """Восстанавливает тензор напряжений без деления касательных компонент."""
    a,b,c,d,e,f=s
    return np.array([[a,d,e],[d,b,f],[e,f,c]])
proof=[]; previous=None
for i in range(1,26):
    row=stresses[i]; sigma=np.array(row[4:10],float); kt,kc=map(float,row[10:12])
    # AcceptedPointDamage выдаёт доли энергии CDP, а не факторы деградации жёсткости.
    dt=kt
    damage_peak=.423*(1-30/(.002*E)); plastic_peak=.002-30/((1-damage_peak)*E)
    ascending=(16+30)*plastic_peak*.5; energy=.64
    if kc*energy<ascending:
        h=(30-16)/plastic_peak
        compression_plastic=2*kc*energy/(16+np.sqrt(16**2+2*h*kc*energy))
    else:
        compression_plastic=plastic_peak-np.log((1-kc)*energy/(energy-ascending))/(30/(energy-ascending))
    dc=-np.expm1(np.log1p(-damage_peak)*compression_plastic/plastic_peak)
    eig=np.linalg.eigvalsh(TestTensor(sigma)); fraction=sum(np.maximum(eig,0))/sum(abs(eig))
    strain=operator@frames[i]; plastic=strain-np.linalg.solve(D,sigma/((1-fraction*dt)*(1-dc)))
    item=dict(state=i,plastic_strain=plastic.tolist(),tension_damage=dt,compression_damage=float(dc),tension_energy_ratio=kt,compression_energy_ratio=kc)
    if previous is not None:item['plastic_change_max']=float(max(abs(plastic-previous)))
    previous=plastic; proof.append(item)
out=repo/'docs/reports/figures/MaterialState_2026-11-26'
data=json.loads((out/'MaterialDiagramData.json').read_bytes())
assert [p['state'] for p in data['concrete']['trajectory']]==list(range(1,26))
err=max(p['plastic_change_max'] for p in proof[15:24])
assert err<1e-12,err
result=dict(ordered_states=list(range(1,26)),plastic_changes=proof,
    states_15_through_24_max_plastic_change=err,constant_damage_15_through_24=True,
    strict_elastic_reconstruction_tolerance=1e-12,strict_elastic_reconstruction_passed=bool(err<1e-12),
    remaining_check='The local elastic unloading/reloading reconstruction is verified; correctness and uniqueness of the full continuation branch remain outside this check.',
    unloading_17_to_18_lambda=[data['concrete']['trajectory'][i-1]['lambda_value'] for i in (17,18)],
    reloading_18_to_24_lambda=[data['concrete']['trajectory'][i-1]['lambda_value'] for i in (18,24)],
    interpretation='Material-point unloading and reloading in accepted continuation; not a uniaxial envelope or proof of branch correctness.',
    source_sha256=sha256((root/'Native/EnvelopeEnergy.Accepted.csv').read_bytes()).hexdigest().upper())
(out/'HookTrajectoryProof.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
(out/'Check-TestHookTrajectory.py').write_bytes(Path(__file__).read_bytes())
print(json.dumps({k:v for k,v in result.items() if k!='plastic_changes'},ensure_ascii=False,indent=2))
